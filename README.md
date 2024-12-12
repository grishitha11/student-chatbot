# Student Details Chatbot

This project is a graphical chatbot interface designed to collect student details interactively and save them in an Excel file. The chatbot guides users step-by-step through a conversational interface, asking for details like Name, Roll Number, Department, Email, and Phone Number.

## Features

- **Graphical User Interface (GUI)**: Built using `tkinter`, the chatbot provides an intuitive and user-friendly interface.
- **Step-by-Step Interaction**: Collects data conversationally, making it engaging for users.
- **Excel Integration**: Saves all collected details into an Excel file named `Student_Details.xlsx` using the `openpyxl` library.
- **Error Handling**: Handles file permissions and ensures smooth operation.

## Prerequisites

- Python 3.6+
- Required Libraries:
  - `openpyxl`
  - `tkinter` (comes pre-installed with Python)

## Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/your_username/student-details-chatbot.git
   cd student-details-chatbot
   ```

2. Install the required Python library:
   ```bash
   pip install openpyxl
   ```

3. Run the chatbot script:
   ```bash
   python chatbot.py
   ```

## How to Use

1. Launch the script using the command above.
2. The chatbot will greet you and ask for your Name.
3. Provide the required details step-by-step:
   - Name
   - Roll Number
   - Department
   - Email
   - Phone Number
4. Once all details are provided, the chatbot will save them to `Student_Details.xlsx`.
5. You can restart the conversation to add details for another student.

## File Structure

- `chatbot.py`: Main script containing the chatbot logic.
- `Student_Details.xlsx`: Excel file where all student details are saved (created automatically).

## Project Details

- **Technologies Used**:
  - Python
  - Tkinter for GUI
  - Openpyxl for Excel file handling

## Contribution

Contributions are welcome! To contribute:

1. Fork the repository.
2. Create a new branch:
   ```bash
   git checkout -b feature-branch-name
   ```
3. Commit your changes and push:
   ```bash
   git commit -m "Description of changes"
   git push origin feature-branch-name
   ```
4. Open a Pull Request.



