# 🛡️ **SnafflerShieldX** - Security Alert Analysis and Management


## 🚀 Overview
 <img width="184" alt="icon" src="https://github.com/user-attachments/assets/a5b9b2d6-28c0-42d1-b86f-eea993df51d0" />



**SnafflerShieldX** is a powerful tool designed to analyze systems for sensitive files and generate detailed reports. It integrates with a modern graphical user interface (GUI) to ease interaction with the user. With a robust scanning engine and a parser to analyze results, **SnafflerShieldX** allows you to visualize security alerts efficiently and clearly.

## 📦 Features

- **System Analysis**: Identifies sensitive files and generates reports.
- **Dynamic Exclusions**: Adds exclusions for removable drives and custom paths.
- **Graphical User Interface**: Intuitive interface built with PyQt6 for easy execution and result display.
- **Detailed Reports**: Generates CSV and JSON reports for detected alerts.
- **Result Visualization**: Displays an alert histogram by severity level and a pie chart for a clear distribution.

### Example GUI
<img width="452" alt="GUI" src="https://github.com/user-attachments/assets/c756b1fd-d11e-4366-b05f-8c8921f17ce9" />

_Example of SnafflerShieldX GUI showing machine details and scan options._

## 📊 Graphical Features

- **Alert visualization by severity level** in a histogram.
- **Alert distribution** by severity level in a pie chart.
- **Directly open CSV reports** through a button in the interface.

### Example Histogram and  Pie Chart
<img width="956" alt="Dashboard" src="https://github.com/user-attachments/assets/908c524c-a59d-436c-83ae-8954aaff207d" />

 
_Visual representation of severity levels from the scan output._



## ⚙️ Prerequisites

Before running **SnafflerShieldX**, make sure the following are installed on your machine:

- **Windows 10 or newer**.
- **Python 3.x** with `pip` installed.
- **WinRAR** installed to decompress files.
- **PowerShell** to execute automation scripts.

## 🔧 Installation

1. **Clone the repository**:
   ```bash
      git clone https://github.com/your-username/SnafflerShieldX.git
      cd SnafflerShieldX
   ```

2. **Start the Setup**:
   ```bash
      Setup.ps1
   ```

## 🖥️ Usage

### Run the program

1. Run the requirments .bat :

   ```bash
      requirments.bat
   ```
2.Run the GUI interface main.pyw
   ```bash
      python main.pyw
   ```

3. **Interface**:
   - Click on **"🚀 Run Snaffler"** to start the scan.
   - View results in the text area.
   - You can display the histogram or open the generated CSV report.

### Main Features:

- **Background scanning**: The scan process runs in the background with real-time output displayed in the interface.
- **Error Handling**: If an error occurs during execution, the program will prompt the user to press **Enter** to resume the scan.
- **Report Display**: Once the scan is complete, you can view the results graphically or open the CSV file.

## 📑 Example of Usage

Once the script is launched, you will see the detected machine name and the option to start the scan process. You can also view the reports as graphical charts (Histogram and Pie Chart).

- The process will scan the system and generate a report detailing security alerts.
- You can then view the histogram for the alert distribution by severity level.


## 🔄 Execution Process

1. The script checks for administrative rights to ensure it has the necessary privileges.
2. It adds exclusions for removable drives and the system disk.
3. It decompresses the `Packages.rar` file to extract the necessary executables.
4. It launches **Snaffler.exe** to perform the scan and generates a JSON file containing the detected alerts.
5. The JSON file is then processed by a PowerShell script to generate a CSV report.
6. The interface displays the results and allows you to open the CSV report or visualize the alerts in graphical form.

## 🛠️ Development

To contribute to the project, follow these steps:

1. Fork this repository.
2. Create a branch for your changes (`git checkout -b feature/feature-name`).
3. Commit your changes (`git commit -am 'Added new feature'`).
4. Push your branch (`git push origin feature/feature-name`).
5. Open a Pull Request for review.

## 🔒 License

This project is licensed under the **MIT License**.

## 💬 Support

For any questions or suggestions, feel free to open an **issue** or contact me directly via GitHub.

---

### Notes:
- Replace `path_to_your_image/gui_screenshot.png` and other placeholders with the actual path to your images within the project repository (or use external image URLs if they are hosted online).
- You can add as many image placeholders as needed throughout the README to visually demonstrate your project's features.

This updated version reflects your new project name, **SnafflerShieldX**, and includes sections where you can showcase images for a more visual and engaging README.
