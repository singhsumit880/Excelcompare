# Excel Compare 📊🔍

A premium, lightweight, and **100% client-side** spreadsheet comparison tool built to run entirely inside the browser. No spreadsheets are ever uploaded to a server, ensuring absolute privacy for financial ledgers, GST reports, and secure client databases.

🔗 **Live Demo:** [singhsumit880.github.io/Excelcompare/](https://singhsumit880.github.io/Excelcompare/)

---

## Key Features 🚀

*   **Tab-by-Position Matching:** Pairs sheets strictly by tab order (first tab vs. first tab, etc.), rendering renamed tabs with clean strikethrough visuals (e.g., `~~Sheet1~~ ➔ SheetA`).
*   **Cell-by-Coordinate Comparison:** Performs precise coordinate-based matching. Standardizes spaces, letter case, decimal precision (e.g., `1.5` vs. `1.50`), and date formatting.
*   **Side-by-Side Workbook Viewer:** Offers a synchronized scroll viewer to compare columns and rows directly next to each other, highlighting changes in yellow, additions in green, and deletions in red.
*   **Smart Search & Filters:** Instantly toggle between:
    *   *Differences Only* (Default)
    *   *Added Rows*
    *   *Missing Rows*
    *   *Modified Values*
*   **Web Worker Architecture:** Leverages multi-threading (background workers) for processing large spreadsheets fluidly without locking the main browser UI thread.
*   **Premium Export Options:**
    *   **Excel (.xlsx):** Generates a multi-tab report mapping overall workbook metrics, tab structural mapping, and precise cell coordinate discrepancies.
    *   **PDF (.pdf):** Outputs a clean auto-table document containing overall statistics and the first 150 cell differences.
    *   **CSV (.csv):** Standard comma-separated values log for quick automated review.

---

## Privacy & Security 🛡️

This tool operates **entirely in your browser**. 
*   No database connections.
*   No API backend or logging.
*   All spreadsheet reading (via `SheetJS`) and difference calculation occur locally on your machine.
*   Fully safe for highly sensitive business files and corporate records.

---

## Tech Stack 🛠️

*   **HTML5 & CSS3:** Semantic markup with variable-driven light/dark themes.
*   **Vanilla JS (ES6):** Client-side state coordination and difference calculations.
*   **Libraries (included locally in `/lib`):**
    *   [SheetJS (xlsx)](https://sheetjs.com/) - Parsing Excel files.
    *   [jsPDF & AutoTable](https://github.com/parallax/jsPDF) - Generating PDF reports.
    *   [Lucide Icons](https://lucide.dev/) - Modern UI icons.

---

## Local Usage 💻

To run the application locally on your computer:

1. Clone or download the repository.
2. Double-click `index.html` to open it directly in any web browser.
*(Note: Browsers restrict Web Workers when running pages on the `file://` protocol. The application automatically detects this and runs the comparison synchronously on the main thread as a fallback, keeping all features fully functional).*

---

## License 📄

This project is open-source and free to use.
