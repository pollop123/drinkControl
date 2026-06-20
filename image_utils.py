import os

from PIL import Image, ImageOps


def compress_image(image_path, max_size=768, quality=88):
    """Resize an image in place while preserving its aspect ratio."""
    if max_size <= 0:
        raise ValueError("max_size must be greater than zero")
    if not 1 <= quality <= 95:
        raise ValueError("quality must be between 1 and 95")

    temp_path = f"{image_path}.compressed.jpg"
    try:
        with Image.open(image_path) as image:
            image = ImageOps.exif_transpose(image)
            image.thumbnail((max_size, max_size), Image.Resampling.LANCZOS)

            if image.mode != "RGB":
                background = Image.new("RGB", image.size, "white")
                if "A" in image.getbands():
                    background.paste(image, mask=image.getchannel("A"))
                else:
                    background.paste(image)
                image = background

            image.save(temp_path, "JPEG", quality=quality, optimize=True)

        os.replace(temp_path, image_path)
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)

    return image_path
