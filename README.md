# Audio Patchlist Generator for Google Docs

A Google Apps Script add-on to automatically generate clean and structured audio patchlists and input/output tables in Google Docs. This tool is perfect for sound engineers and event planners.

---

## 🚀 Features

-   **Fast & Efficient:** Generates clean, structured tables in seconds.
-   **Customizable:** Choose which sections (**Console**, **StageBox**, **Aux**) you need and specify the number of rows for each table.
-   **Organized Layout:** Each table is placed on a new page, neatly grouped by category (**Inputs**, **Outputs**).
-   **Clear Labels:** Tables include clear headings with an optional console name for easy identification.
-   **Automatic Numbering:** The first column of each table is automatically numbered, starting from one.

---

## 🛠️ How to Install and Use

Since this is a Google Apps Script project, you will use it directly within your Google Docs environment.

### 1. Copy the Script
1.  Open a new Google Docs document.
2.  Go to **Extensions > Apps Script**.
3.  Delete any existing code and copy the content of the following files into your project:
    -   **`Code.gs`**: The main script logic.
    -   **`Sidebar.html`**: The code for the user interface of the sidebar.
    
### 2. Run the Script
1.  Save the files in the Apps Script editor.
2.  Return to your Google Docs document.
3.  Refresh the page.
4.  You will now see a new menu item called **Tabel Generator** appear under the **Extensions** menu.
5.  Click on **Extensions > Tabel Generator > Open Generator** to open the sidebar.

### 3. Generate Tables
1.  In the sidebar, you can enter the **Console Name**.
2.  Toggle on the desired sections (**Console**, **StageBox**, **Aux**).
3.  Enter the required number of **Input** and **Output** rows.
4.  Click the **OK** button to generate the tables.

---

## 📄 License

This project is licensed under the **MIT License**. See the `LICENSE` file for more details.

---

## 👨‍💻 Contributing

Feel free to contribute to this project. If you find a bug or want to add a new feature, you can open an **issue** or submit a **pull request**.
