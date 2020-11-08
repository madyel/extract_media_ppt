from pptx import Presentation
from pptx.enum.shapes import PP_MEDIA_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from zipfile import ZipFile
import os

DEFAULT_PATH_PPT = 'ppt/media'

class Ext:
    def __init__(self):
        self._ext_permission_video = ['mp4', 'avi', 'mpg', 'mpeg', 'wmv']
        self._ext_permission_image = ['png', 'jpeg', 'jpg', 'bmp', 'svg']
    @property
    def ext_permission_video(self):
        return self._ext_permission_video
    @ext_permission_video.setter
    def ext_permission_video(self, ext):
        if len(ext) != 0:
            self._ext_permission_video = ext
    @property
    def ext_permission_image(self):
        return self._ext_permission_image
    @ext_permission_image.setter
    def ext_permission_image(self, ext):
        if len(ext) != 0:
            self._ext_permission_image = ext

class PowerPoint(object):
    def __init__(self, filename, type, output='temp', *file_extension):
        self.ext = Ext()
        self.filename = filename
        self.output = output
        self.type=type
        self.prs = Presentation(self.filename)
        if self.type == 'image':
            self.ext.ext_permission_image = file_extension
        else:
            self.ext.ext_permission_video = file_extension

    def getInfoImage(self,media):
        if hasattr(media, 'shape_type'):
            if media.shape_type == MSO_SHAPE_TYPE.PICTURE:
                info = {}
                self.num_media = self.num_media + 1
                info['shape_id'] = media.shape_id
                info['filename'] = "image" + str(self.num_media)
                info['num_slide'] = self.num_slide
                info['num_media'] = self.num_media
                self.infos.append(info)

    def getInfoVideo(self,media):
        if hasattr(media, 'media_type'):
            if media.media_type == PP_MEDIA_TYPE.MOVIE:
                info = {}
                self.num_media = self.num_media + 1
                info['shape_id'] = media.shape_id
                info['filename'] = "media" + str(self.num_media)
                info['num_slide'] = self.num_slide
                info['num_media'] = self.num_media
                self.infos.append(info)

    def __start(self):
        self.infos = list()
        self.num_slide = 0
        self.num_media = 0
        for slide in self.prs.slides:
            self.num_slide = self.num_slide + 1
            for media in slide.shapes:
                if self.type == 'image':
                    self.getInfoImage(media)
                if self.type == 'video':
                    self.getInfoVideo(media)


    def has_video(self, prs):
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, 'media_type'):
                    if shape.media_type == PP_MEDIA_TYPE.MOVIE:
                        return True
        return False


    def check_idSlide(self, filename, list):
        for ls in list:
            if filename in ls['filename']:
                return ls['num_slide']

    def extractAllMedia(self):
        check_ext = 0
        with ZipFile(self.filename, 'r') as zipObject:
           listOfFileNames = zipObject.namelist()
           for fileName in listOfFileNames:
               if fileName.startswith(DEFAULT_PATH_PPT):
                   tmp = self.output
                   if not os.path.exists(tmp):
                       os.makedirs(tmp)
                   zipObject.extract(fileName, tmp)
                   check_ext = check_ext + 1
        if check_ext == 0:
           print("Not Found!")
        else:
           print(f'Completed, {check_ext} media')



    def extract(self):
        self.__start()
        check_ext = 0
        ext = self.ext.ext_permission_video
        if self.type == 'image':
            ext = self.ext.ext_permission_image
        if len(self.infos) == 0:
            print("Not Found!")
            exit()
        with ZipFile(self.filename, 'r') as zipObject:
           listOfFileNames = zipObject.namelist()
           for fileName in listOfFileNames:
               if fileName.startswith(DEFAULT_PATH_PPT):
                   filename, file_extension = os.path.splitext(fileName)
                   name = os.path.basename(filename)
                   if file_extension[1:] in ext:
                       tmp = self.output +'/'+ str(self.check_idSlide(name, self.infos))
                       if not os.path.exists(tmp):
                           os.makedirs(tmp)
                       zipObject.extract(fileName, tmp)
                       check_ext = check_ext + 1
        if check_ext == 0:
            print("Not Found!")
        else:
            print(f'Completed, {check_ext} media')