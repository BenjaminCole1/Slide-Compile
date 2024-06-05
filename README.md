# Slide-Compile

## A Text-Based Slideshow Editor

### Dependencies
- Windows
- Python 3.7 or higher
- python-pptx
- Pillow
- requests

### Installation Guide

1. **Install Python:**
   - Ensure Python 3.7 or higher is installed on your system. You can download Python from [python.org](https://www.python.org/downloads/).
   - Make sure to tick "pip" and "tcl/tk and IDLE" when installing python. If you have python installed already, download the installer open it and click modify. 

2. **Install Required Packages:**
   - Open a command prompt and run the following commands to install the necessary Python packages:
     ```sh
     pip install python-pptx
     pip install Pillow
     pip install requests
     ```

3. **Download Slide-Compile:**
   - Download the `SlideCompile.py` script and place it in a desired directory on your computer.

### Usage
1. Run `SlideCompile.py`
2. Write slides with the syntax shown in the syntax section
3. Press Compile; if you didn't save, it will prompt you to save to a location, and then it will prompt you to save the pptx (PowerPoint file) to a location on your computer
4. If you get any errors or you can't get it working, first try troubleshooting them within the "Common Errors" section in this readme
5. Edit further in a PowerPoint software to customize more

### Syntax
Syntax is the following; it is not case sensitive

- **NewSlide/New_Slide**: This creates a slide; you specify the contents underneath this. It does not require a number or a colon.
  
- **FormatNumber:/Format_Number:** This specifies the format number. Valid format numbers range from 1 to 10, see valid formats with the `Formats.pptx` file in the folder formats. This is followed by a colon, space, and a number. For example:
  ```
  FormatNumber: 1
  ```
  

- **Title:** This specifies the title content. Not all formats support titles, and the program will show an error and it won't compile if you try to make a title on a format that doesn't have one. Similarly to FormatNumber, its syntax looks like the following:
  ```
  Title: This is a title!
  ```

- **Content:** This specifies the bullet points in the content part of your PowerPoint. Similar to the title, not all formats support titles, and the program will throw an error. Syntax:
  ```
  Content: this is some content with a bullet point in front of it
  ```

- **Image:** This specifies the image URL to be added to the slide. Ensure the image URL is accessible and valid. Syntax:
  ```
  Image: http://example.com/image.png
  ```

- **ImagePosition:/Image_Position:** This specifies the position of the image on the slide. Valid positions are:
  - top left
  - top right
  - bottom left
  - bottom right
  - middle left
  - middle right
  - middle
  Syntax:
  ```
  ImagePosition: top right
  ```

- **ImageSize:/Image_Size:** This specifies the size of the image. Valid sizes are:
  - tiny
  - small
  - medium
  - large
  - extra large
  Syntax:
  ```
  ImageSize: medium
  ```

### Available Formats
You can find the available formats in the `Formats.pptx` file located in the `Formats` folder. This file showcases the different format options you can use with Slide-Compile.

### Common Errors
- **Could not open file:** Make sure that Windows Defender isn't blocking the reading or writing permissions of the program.

- **Could not Save File:** Make sure that Windows Defender isn't blocking the reading or writing permissions of the program.

### Example SlideShow
You can see what this looks like by going downloading the Example.pptx file, and loading it into your favorite SlideShow Editor

```
NewSlide

FormatNumber: 1

Title: The Benefits of Walking Every Day

NewSlide

FormatNumber: 2

Title: Health Benefits #1

Content: Improves cardiovascular health

NewSlide

FormatNumber: 2

Title: Benefits #2

Content: Helps maintain a healthy weight
Content: Strengthens muscles and bones
Content: Enhances mental clarity and creativity

NewSlide

FormatNumber: 3

Title: Adding an Image

Image: http://example.com/image.png
ImagePosition: middle
ImageSize: large
```
