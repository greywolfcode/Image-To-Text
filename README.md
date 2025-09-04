# Image-To-Text
Makes a word document where the font colours match an image's pixels

## Usage
```
imgToText.py imagePath textPath outputPath [-r] [-t]
```
The textPath should be a .txt file


-r, --repeat makes repeats the text until it is the size of or bigger than the image.

-t, --truncate truncates the text if it is larger then the image


Note that the more text there is, the longer it will take to finish teh conversion
