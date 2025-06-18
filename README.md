# OCR Annotation Tool

## Overview
This Python application serves as an OCR annotation tool, allowing users to annotate text extracted from images via Optical Character Recognition (OCR). The tool provides functionalities to open a CSV file containing image filenames and their corresponding labels, open an image folder for annotation, display images, edit labels, navigate between images, and update the CSV file with the annotations.

## Features
- **Open CSV File**: Load a CSV file containing image filenames and labels.
- **Open Image Folder**: Select a folder containing images for annotation.
- **Display Image and Label**: Show the selected image along with its current label.
- **Edit Label**: Update the label for the displayed image.
- **Navigate Images**: Move between images in the folder.
- **Update Annotations**: Save changes to the labels back to the CSV file.
- **Delete Image**: Remove the currently displayed image and its corresponding label.

## Installation
To run the OCR annotation tool, ensure you have Python 3.11.5 installed on your system. Then, follow these steps:
1. Clone the repository containing the Python script.
2. Install the required dependencies using pip:

```bash
pip install -r requirements.txt
```

Sure, here's the markdown documentation for your Python code:
✅ To fix it, you need to install PyInstaller:
Option 1: Install Globally (System-wide)

```bash
sudo apt-get update
sudo apt-get install python3-pip
pip3 install pyinstaller
```

Then verify installation:

```bash
pyinstaller --version
```

Option 2: Install in a Virtual Environment (Recommended)
This keeps your environment clean:

```bash
python3 -m venv venv
source venv/bin/activate
```

```bash
pyinstaller --onefile --windowed app-premium.py
```

# OCR Annotation Tool

## Overview
This Python application serves as an OCR annotation tool, allowing users to annotate text extracted from images via Optical Character Recognition (OCR). The tool provides functionalities to open a CSV file containing image filenames and their corresponding labels, open an image folder for annotation, display images, edit labels, navigate between images, and update the CSV file with the annotations.

## Features
- **Open CSV File**: Load a CSV file containing image filenames and labels.
- **Open Image Folder**: Select a folder containing images for annotation.
- **Display Image and Label**: Show the selected image along with its current label.
- **Edit Label**: Update the label for the displayed image.
- **Navigate Images**: Move between images in the folder.
- **Update Annotations**: Save changes to the labels back to the CSV file.
- **Delete Image**: Remove the currently displayed image and its corresponding label.

## Installation
To run the OCR annotation tool, ensure you have Python 3.11.5 installed on your system. Then, follow these steps:
1. Clone the repository containing the Python script.
2. Install the required dependencies using pip: ```pip install PyQt5```
3. Run the script: ```python app.py```



## Usage
1. **Open CSV File**: Click the "Open CSV File" button to load a CSV file containing image filenames and labels.
2. **Open Image Folder**: Click the "Open Image Folder" button to select a folder containing images for annotation.
3. **Display Image**: The tool displays the first image from the selected folder along with its label.
4. **Edit Label**: Edit the label text in the provided text box.
5. **Navigate Images**: Use the "Next" and "Previous" buttons to move between images in the folder.
6. **Update Annotations**: Click the "Update" button to save changes to the label back to the CSV file.
7. **Delete Image**: Click the "Delete" button to remove the currently displayed image and its label.

## Technologies Used
- **Python**: Programming language used for development.
- **PyQt5**: GUI toolkit for creating desktop applications.
- **CSV Module**: Python module for reading and writing CSV files.
- **os Module**: Python module for interacting with the operating system.
- **QFileDialog**: PyQt5 class for opening file dialogs.
- **QPixmap**: PyQt5 class for displaying images.
- **QMessageBox**: PyQt5 class for displaying message boxes.

## Contributors
- **[Bipin Saha]**: Developer (bipinsaha.bd@gmail.com)
- **[Contributor Name]**: Contributor

## License
This project is licensed under the [MIT License](LICENSE).
