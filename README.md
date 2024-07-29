# AI-TEXT-TO-WORD

## Description

AI-TEXT-TO-WORD is a desktop application that converts formatted text (typically generated by AI) into a well-structured Word document (.docx). The application provides a simple and intuitive graphical interface for pasting text and generating the corresponding Word document.

## Features

- User-friendly graphical interface with PyQt6
- Conversion of formatted text to Word document
- Automatic management of paragraph styles (headings, subheadings, bullet points)
- File existence check before overwriting
- Error handling (permissions, etc.)

## Prerequisites

- Python 3.6+
- PyQt6
- python-docx

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/your-username/AI-TEXT-TO-WORD.git
   ```

2. Navigate to the project directory:
   ```
   cd AI-TEXT-TO-WORD
   ```

3. Install the dependencies:
   ```
   pip install PyQt6 python-docx
   ```

## Usage

1. Launch the application:
   ```
   python main.py
   ```

2. Paste your formatted text into the text area.

3. Click the "Generate Word Document" button.

4. Choose where to save the Word file.

5. If the file already exists, the application will ask if you want to replace it.

## Supported Text Format

The application supports the following text format:

- `###` for level 1 headings
- `####` for level 2 headings
- `- **text**` for level 3 headings
- `-` for bullet point items
- Any other text is considered normal text

## Code Structure

- `TextToWordConverter`: Main application class
  - `initUI()`: Initializes the user interface
  - `generate_word_document()`: Handles Word document generation
  - `convert_to_word()`: Converts text to Word document

## Contributing

Contributions are welcome! Feel free to open an issue or submit a pull request.

## License

[MIT License](https://opensource.org/licenses/MIT)
