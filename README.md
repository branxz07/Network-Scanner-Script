# Network Scanner Script

This Python script scans your local network for devices, logs the details (IP, MAC, Hostname, Vendor/Fabricant) to an Excel file. The script is designed to run on Linux and Windows systems.

## Features

- Scans the local network.
- Retrieves the hostname and vendor information for each device.
- Logs the device information to an Excel file.
- Automatically adjusts column widths and adds filters to the Excel file.

## Requirements

- Python 3.x
- `requests` library
- `openpyxl` library
- `pandas` library

## Installation

1. **Clone the repository:**

    ```sh
    git clone https://github.com/branxz07/Network-Scanner-Script.git
    cd Network-Scanner-Script
    ```

2. **Install the required Python packages:**

    ```sh
    pip install requests openpyxl pandas
    ```

## Usage

1. **Run the script:**

    ```sh
    python network_scanner.py
    ```

    The script will identify the operating system and begin scanning the local network using the adecuate tool depending on the OS, in case of windows will be using the ARP table. It will log the device information to an Excel file named `network-log-YYYY-MM-DD.xlsx`, where `YYYY-MM-DD` is the current date.

2. **Log file:**

    The log file will be created in the same directory as the script. Each scan will append new data to the log file, including the timestamp of the scan.

## Notes
### For Windows
- Ensure that the `arp` command is available on your system. This command is typically available by default on most Linux and Windows systems.
- The script makes HTTP requests to `https://api.macvendors.com` to get MAC vendor information. Ensure that your system has internet access.
- If the log file is open in another program (e.g., Excel), the script will retry saving the log file every 5 seconds until it succeeds.
### For Linux

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any improvements or bug fixes.

## License

This project is free to use. You are free to use, modify, and distribute this code without any restrictions.

## Disclaimer

This project is provided "as is". Users of this code are responsible for understanding its functionality and implementing it appropriately in their own projects. While efforts have been made to ensure its reliability, users should exercise due diligence in testing and adapting the code to their specific needs. The authors and contributors are not liable for any outcomes resulting from the use of this code.

## Contact

For any questions or suggestions, please open an issue.
