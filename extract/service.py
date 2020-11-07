from pptx import Presentation
from pptx.enum.shapes import PP_MEDIA_TYPE
from zipfile import ZipFile
import os
from pptx.util import Emu

DEFAULT_PATH_PPT = 'ppt/media'

class Ext:
    def __init__(self):
        self._ext_permission = ['mp4', 'avi', 'mpg', 'mpeg', 'wmv']
    @property
    def ext_permission(self):
        return self._ext_permission
    @ext_permission.setter
    def ext_permission(self, ext):
        if len(ext) != 0:
            self._ext_permission = ext


class PowerPoint(object):
    def __init__(self, filename, output='temp', *file_extension):
        self.ext = Ext()
        self.filename = filename
        self.output = output
        self.prs = Presentation(self.filename)
        self.ext.ext_permission = file_extension

    def __start(self):
        self.infos = list()
        num_slide = 0
        w_emu = Emu(self.prs.slide_width)
        h_emu = Emu(self.prs.slide_height)

        num_media = 0
        for slide in self.prs.slides:
            num_slide = num_slide + 1
            for movie in slide.shapes:
                if hasattr(movie, 'media_type'):
                    if movie.media_type == PP_MEDIA_TYPE.MOVIE:
                        info = {}
                        num_media = num_media + 1

                        info['shape_id'] = movie.shape_id
                        info['filename'] = "media"+str(num_media)
                        info['num_slide'] = num_slide
                        info['num_media'] = num_media

                        x = Emu(movie._element.x)
                        y = Emu(movie._element.y)
                        cx = Emu(movie._element.cx)
                        cy = Emu(movie._element.cy)

                        w = w_emu.inches*96
                        h = h_emu.inches*96
                        info['x'] = round(float(x.inches*96/w),3)
                        info['y'] = round(float(y.inches*96/h),3)
                        info['cx'] = round(float(cx.inches*96/w),3)
                        info['cy'] = round(float(cy.inches*96/h),3)
                        ##yield movie
                        self.infos.append(info)

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


    def extractVideo(self):
        self.__start()
        with ZipFile(self.filename, 'r') as zipObject:
           listOfFileNames = zipObject.namelist()
           for fileName in listOfFileNames:
               if fileName.startswith(DEFAULT_PATH_PPT):
                   filename, file_extension = os.path.splitext(fileName)
                   name = os.path.basename(filename)
                   if file_extension[1:] in self.ext.ext_permission:
                       tmp = self.output +'/'+ str(self.check_idSlide(name, self.infos))
                       if not os.path.exists(tmp):
                           os.makedirs(tmp)
                       zipObject.extract(fileName, tmp)
