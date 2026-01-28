# Image-To-Text
Makes a word document where the font colours match an image's pixels.

## Usage
```
imgToText.py imagePath textPath outputPath [-r] [-t]
```
The textPath should be a .txt file


-r, --repeat :  repeats the text until the length of text is the size of or bigger than the number of pixels in the image.

-t, --truncate :  truncates the text if the text length is longer then the number of pixels in the image.


Note that the more text there is, the longer it will take to finish the conversion.

The output document will probably not look right, as the normal text size is often too large. Shrinking the text size should fix it.

## Dependencies
- [Pillow](https://github.com/python-pillow/Pillow)
- [python-docx](https://github.com/python-openxml/python-docx)
