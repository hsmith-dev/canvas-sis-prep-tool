# Canvas SIS Prep Tool

[![Latest Release](https://img.shields.io/github/v/release/hsmith-dev/canvas-sis-prep-tool?style=for-the-badge)](https://github.com/hsmith-dev/canvas-sis-prep-tool/releases)
[![MIT License](https://img.shields.io/github/license/hsmith-dev/canvas-sis-prep-tool?style=for-the-badge)](https://github.com/hsmith-dev/canvas-sis-prep-tool/blob/main/LICENSE)
[![Issues](https://img.shields.io/github/issues/hsmith-dev/canvas-sis-prep-tool?style=for-the-badge)](https://github.com/hsmith-dev/canvas-sis-prep-tool/issues)

A user-friendly desktop application designed to streamline the creation of CSV files for Canvas SIS imports, with a focus on simplifying course shell creation.

---
![Screenshot of the Canvas SIS Prep Tool application](https://github.com/hsmith-dev/canvas-sis-prep-tool/blob/main/app-image.png?raw=true)

## 🎯 About The Project

For educational administrators, instructional designers, and anyone responsible for bulk course creation in Canvas, manually building SIS import files is a tedious and error-prone process. This tool provides an intuitive graphical interface to manage all the required data points—people, courses, terms, and enrollments—and generates perfectly formatted `courses.csv`, `sections.csv`, and `enrollments.csv` files ready for upload.

## ✨ Key Features

* **Intuitive Graphical Interface:** A clean, tabbed UI for managing all your data in one place.
* **Centralized Data Management:** Easily add, edit, and delete records for:
    * 👥 People (Students, Teachers, etc.)
    * 📚 Courses
    * 🏢 Program Area
    * 🗓️ Terms
    * 📂 Accounts
* **Effortless Section & Enrollment Creation:** A guided process to create course sections and enroll users with just a few clicks.
* **One-Click CSV Generation:** Automatically produce the three essential files (`courses.csv`, `sections.csv`, `enrollments.csv`) with the correct headers and relationships.
* **Data Portability:** Import and export your core data lists (people, courses, etc.) to and from CSVs, making it easy to back up your work or migrate between machines.
* **Customizable Theme:** Toggle between a light and dark mode for your viewing comfort.

## 🚀 Getting Started

No installation is required! The application is a portable executable.

1.  Navigate to the [**Releases Page**](https://github.com/hsmith-dev/canvas-sis-prep-tool/releases).
2.  Download the `CanvasSISPrepTool.exe` or `CanvasSISPrepTool.app` file from the latest release's **Assets**.
3.  Run the executable file.

## 📋 Usage Walkthrough

The intended workflow is straightforward:

1.  **Populate Core Data:** Start by adding your institutional data in the `People`, `Courses`, `Program Area`, `Terms`, and `Accounts` tabs. You can also use the `Import` feature on the `Actions` tab.
2.  **Create a Section:** Navigate to the `Sections & Enrollments` tab and click "Create Section". Use the dialog to link a Course, Term, and Account to create a unique course section.
3.  **Add Enrollments:** With a section selected, click "Manage Enrollments" to add people (students, teachers) to that specific section with their designated roles.
4.  **Generate CSVs:** Once your sections and enrollments are configured, go to the `Actions` tab, click "Generate Canvas CSV Files...", choose a save location, and the tool will export the formatted files for you.

## 🛠️ Built With

* [Python 3](https://www.python.org/)
* [Tkinter](https://docs.python.org/3/library/tkinter.html) (for the graphical user interface)
* [PyInstaller](https://pyinstaller.org/) (for packaging into an executable)

## 🤝 Contributing

Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply [open an issue](https://github.com/hsmith-dev/canvas-sis-prep-tool/issues) with the tag "enhancement".

Don't forget to give the project a star! Thanks again!

## 📄 License

Distributed under the GPL-3 License. See `LICENSE` for more information.

## 📧 Contact

Harrison Smith - harrison@hsmith.dev

Website Link: [https://harrisonsmith.ai/canvas-sis-import-tool](https://harrisonsmith.ai/canvas-sis-import-tool)
