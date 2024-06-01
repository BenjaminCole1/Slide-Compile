# Slide-Compile
A Text-Based Slideshow Editor

## Dependencies
Windows, Python, Python-pptx

## Usage
1) Run SlideCompile.py
2) Write slides with the syntax shown in the syntax section
3) Press Compile, if you didn't save, it will prompt you to save to a location, and then it will prompt you to save the pptx (powerpoint file) to a location on your computer
4) if you get any errors or you can't get it working, first try troubleshooting them within the "Common Errors" section in this readme
5) edit further in a powerpoint software to customize more
____________________________________________________________________________________________
## Syntax
Syntax is the following, it is not case sensitive

### NewSlide/New_Slide
This creates a slide, you specify the contents underneath this. It does not require a number or a colon.

### FormatNumber:/Format_Number:
This specifies the format number. Valid format numbers range from 1 to 10, see valid formats with the Formats.pptx file in the folder formats
This is followed by a colon, space, number. For example

FormatNumber: 1


### Title:
This specifies the title content. Not all formats support titles, and the program will show an error and it won't compile if you try to make a title on a format that doesn't have one.
Similarly to FormatNumber, its syntax looks like the following:

Title: This is a title!


### Content:
This specifies the bullet points in the content part of your powerpoint. similar to the title, not all formats support titles, and the program will throw an error.
Syntax:

Content: this is some content with a bullet point in front of it

___________________________
## Common Errors
### Could not open file
make sure that windows defender isn't blocking the reading or writing permissions of the program

### Could not Save File
make sure that windows defender isn't blocking the reading or writing permissions of the program

# Example SlideShow
## see this example in Example.pptx

NewSlide

FormatNumber: 1

Title: The Benefits of Walking Every Day

NewSlide

FormatNumber: 2

Title: Health benifits #1

Content: Improves cardiovascular health

NewSlide

FormatNumber: 2

Title: Benifits #2

Content: Helps maintain a healthy weight

Content: Strengthens muscles and bones

Content: Enhances mental clarity and creativity
