# Extract Media PowerPoint

[![PyPI version](https://badge.fury.io/py/Extract-Media-PowerPoint.svg)](https://badge.fury.io/py/Extract-Media-PowerPoint)
[![Python](https://img.shields.io/pypi/pyversions/Extract-Media-PowerPoint.svg)](https://pypi.org/project/Extract-Media-PowerPoint/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE.txt)

A lightweight Python library to extract images and videos from PowerPoint (`.pptx`) presentations, with optional filtering by media type, file extension, and automatic organization by slide number.

## Features

- Extract **all** embedded media (images and videos) from a `.pptx` file
- Extract **filtered** media by type (`image` or `video`) and custom extensions
- Output organized in subdirectories by **slide number**
- Supports custom output directories
- Built on top of [`python-pptx`](https://python-pptx.readthedocs.io/)

## Requirements

- Python >= 3.10
- lxml >= 6.0.2
- Pillow >= 12.1.1
- python-pptx >= 1.0.2

## Installation

```bash
pip install Extract-Media-PowerPoint
```

## Quick Start

### Extract all media (no filtering)

```python
from extract import PowerPointMediaExtractor

extractor = PowerPointMediaExtractor(
    filepath="presentation.pptx",
    output_dir="output"
)
count = extractor.extract_all_media()
print(f"Extracted {count} media files")
```

Files are saved under `output/ppt/media/`.

### Extract images filtered by extension

```python
from extract import PowerPointMediaExtractor

extractor = PowerPointMediaExtractor(
    filepath="presentation.pptx",
    media_type="image",
    output_dir="output",
    extensions=["png", "jpg"]
)
count = extractor.extract_filtered_media()
```

Files are saved under `output/<slide_number>/ppt/media/`.

### Extract videos

```python
from extract import PowerPointMediaExtractor

extractor = PowerPointMediaExtractor(
    filepath="presentation.pptx",
    media_type="video",
    output_dir="output",
    extensions=["mp4", "avi"]
)
count = extractor.extract_filtered_media()
```

## API Reference

### `PowerPointMediaExtractor`

```python
PowerPointMediaExtractor(
    filepath: str | Path,
    media_type: str = "image",      # "image" or "video"
    output_dir: str | Path = "temp",
    extensions: list[str] | None = None,
)
```

| Parameter    | Type                    | Default    | Description                                              |
|-------------|-------------------------|------------|----------------------------------------------------------|
| `filepath`  | `str \| Path`           | —          | Path to the `.pptx` file                                |
| `media_type`| `str`                   | `"image"`  | Media type to extract: `"image"` or `"video"`           |
| `output_dir`| `str \| Path`           | `"temp"`   | Directory where media will be saved                      |
| `extensions`| `list[str] \| None`     | `None`     | Allowed extensions (defaults to all for the media type) |

**Default extensions:**

| Type    | Default extensions                    |
|---------|---------------------------------------|
| `image` | `png`, `jpeg`, `jpg`, `bmp`, `svg`   |
| `video` | `mp4`, `avi`, `mpg`, `mpeg`, `wmv`   |

#### Methods

| Method                    | Returns | Description                                              |
|--------------------------|---------|----------------------------------------------------------|
| `extract_all_media()`    | `int`   | Extracts all embedded media, returns count               |
| `extract_filtered_media()` | `int` | Extracts media filtered by type/extension, organized by slide |

### `MediaInfo`

Dataclass representing a media item found in the presentation.

```python
@dataclass
class MediaInfo:
    shape_id: int
    filename: str
    slide_number: int
```

## Logging

The library uses Python's standard `logging` module under the logger name `extract.service`. To see output:

```python
import logging
logging.basicConfig(level=logging.INFO)
```

## License

MIT — see [LICENSE.txt](LICENSE.txt).
