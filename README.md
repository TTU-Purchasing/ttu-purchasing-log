# README.md

# TTU Purchase Orders Log

This repository contains the source code for a **Streamlit application** designed to process and analyze purchase order logs. The app provides visual insights into key performance indicators and generates customizable reports.

---

## Features
- Upload and process Excel files with purchase order data.
- Apply filters based on dates, requisitioners, and purchase accounts.
- Visualize on-time delivery performance, top vendors, and top items ordered.
- Generate PDF reports with key metrics and visualizations.

---

## Requirements

### Install Dependencies
1. Python 3.8 or later.
2. Install required Python libraries by running:
   ```bash
   pip install -r requirements.txt
   ```

### File Requirements
The application expects an Excel file containing the following columns:
- **OrderDate**
- **PONumber**
- **Total**
- **VendorName**
- **Requisitioner**
- **Purchase Account** (optional)
- **RecDate** (optional)
- **RequestDate** (optional)

Ensure that your file contains these columns before uploading it to the application.

---

## How to Run the Application
1. Clone the repository:
   ```bash
   git clone https://github.com/your-repo-name.git
   cd your-repo-name
   ```
2. Install dependencies as shown above.
3. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```
4. Open the URL provided by Streamlit to access the application in your browser.

---

## Customization
### Adding a Logo
- Replace the `TTU_LOGO.jpg` file with your organization's logo.

### Styling
- Update the CSS in `app.py` to customize the app's appearance.

---

## Contributing
1. Fork the repository and create a new branch for your feature or bugfix.
2. Submit a pull request with a clear description of your changes.

---

## Notes
- Ensure no sensitive or internal information is included when uploading files.
- This repository does not include any actual purchase order data.

---

## License
[Specify your license here, e.g., MIT License]
