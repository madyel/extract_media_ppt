import os
from zipfile import ZipFile
from pptx import Presentation
from pptx.enum.shapes import PP_MEDIA_TYPE, MSO_SHAPE_TYPE

DEFAULT_PPT_MEDIA_PATH = 'ppt/media'

class AllowedExtensions:
    def __init__(self):
        self._video = ['mp4', 'avi', 'mpg', 'mpeg', 'wmv']
        self._image = ['png', 'jpeg', 'jpg', 'bmp', 'svg']

    @property
    def video(self):
        return self._video

    @video.setter
    def video(self, ext_list):
        if ext_list:
            self._video = list(ext_list)

    @property
    def image(self):
        return self._image

    @image.setter
    def image(self, ext_list):
        if ext_list:
            self._image = list(ext_list)

class PowerPointMediaExtractor:
    def __init__(self, filepath, media_type='image', output_dir='temp', *extensions):
        self.extensions = AllowedExtensions()
        self.filepath = filepath
        self.output_dir = output_dir
        self.media_type = media_type.lower()
        self.presentation = Presentation(self.filepath)

        if self.media_type == 'image':
            self.extensions.image = extensions
        elif self.media_type == 'video':
            self.extensions.video = extensions
        else:
            raise ValueError("Invalid media_type. Use 'image' or 'video'.")

        self._infos = []
        self._current_slide_num = 0
        self._media_counter = 0

    def _collect_media_info(self):
        self._infos = []
        self._media_counter = 0

        for idx, slide in enumerate(self.presentation.slides, start=1):
            self._current_slide_num = idx
            for shape in slide.shapes:
                if self.media_type == 'image':
                    self._extract_image_info(shape)
                elif self.media_type == 'video':
                    self._extract_video_info(shape)

    def _extract_image_info(self, shape):
        if getattr(shape, 'shape_type', None) == MSO_SHAPE_TYPE.PICTURE:
            self._media_counter += 1
            self._infos.append({
                'shape_id': shape.shape_id,
                'filename': f"image{self._media_counter}",
                'slide_number': self._current_slide_num
            })

    def _extract_video_info(self, shape):
        if getattr(shape, 'media_type', None) == PP_MEDIA_TYPE.MOVIE:
            self._media_counter += 1
            self._infos.append({
                'shape_id': shape.shape_id,
                'filename': f"media{self._media_counter}",
                'slide_number': self._current_slide_num
            })

    def _find_slide_for_filename(self, filename_stem):
        for item in self._infos:
            if filename_stem in item['filename']:
                return item['slide_number']
        return None

    def extract_all_media(self):
        extracted = 0
        with ZipFile(self.filepath, 'r') as archive:
            for name in archive.namelist():
                if name.startswith(DEFAULT_PPT_MEDIA_PATH):
                    os.makedirs(self.output_dir, exist_ok=True)
                    archive.extract(name, self.output_dir)
                    extracted += 1

        print(f"{'Completed' if extracted else 'Not Found!'}, {extracted} media")

    def extract_filtered_media(self):
        self._collect_media_info()
        if not self._infos:
            print("No media found.")
            return

        allowed_exts = self.extensions.image if self.media_type == 'image' else self.extensions.video
        allowed_exts = [e.lower() for e in allowed_exts]
        extracted = 0

        with ZipFile(self.filepath, 'r') as archive:
            for name in archive.namelist():
                if name.startswith(DEFAULT_PPT_MEDIA_PATH):
                    filename, file_ext = os.path.splitext(name)
                    file_ext = file_ext[1:].lower()
                    if file_ext in allowed_exts:
                        base_name = os.path.basename(filename)
                        slide_number = self._find_slide_for_filename(base_name)
                        if slide_number is not None:
                            target_path = os.path.join(self.output_dir, str(slide_number))
                            os.makedirs(target_path, exist_ok=True)
                            archive.extract(name, target_path)
                            extracted += 1

        print(f"{'Completed' if extracted else 'Not Found!'}, {extracted} media")
