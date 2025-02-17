from PIL import Image
import numpy as np

# ASCII characters used for grayscale mapping
ASCII_CHARS = ["@", "#", "S", "%", "?", "*", "+", ";", ":", ",", "."]

def resize_image(image, new_width=100):
    width, height = image.size
    aspect_ratio = height / width / 1.65  # Adjust for character width
    new_height = int(new_width * aspect_ratio)
    return image.resize((new_width, new_height))

def grayscale(image):
    return image.convert("L")

def pixels_to_ascii(image):
    pixels = np.array(image)
    ascii_str = ""
    for row in pixels:
        for pixel in row:
            ascii_str += ASCII_CHARS[pixel // 25]  # map pixel to ASCII
        ascii_str += "\n"
    return ascii_str

def main(image_path):
    try:
        image = Image.open(image_path)
    except Exception as e:
        print(e)
        return

    image = resize_image(image)
    image = grayscale(image)
    ascii_str = pixels_to_ascii(image)

    with open("ascii_image.txt", "w") as f:
        f.write(ascii_str)

    print(ascii_str)

# Replace 'image.jpg' with your image path
main("D:\\Downloads\\Images\\zTmAOY2-light-yagami-wallpaper2.jpg")


