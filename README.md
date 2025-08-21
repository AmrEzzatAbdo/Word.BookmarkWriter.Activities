# 📄 Word.BookmarkWriter.Activities

A **UiPath custom activity package** designed for automation developers who work with Microsoft Word templates.  
This package enables you to **fill Word bookmarks dynamically with values from a list** and optionally **export the document to PDF** in the same output path.

---

## 🚀 Features

- 🔖 **Insert Data into Word Bookmarks**  
  Provide key-value pairs and automatically fill Word bookmarks.

- 📂 **Flexible Input**  
  Accepts lists or dictionaries of values for bulk insertion.

- 📝 **Non-UI Based**  
  Runs in background automation without requiring UI automation of Word.

- 📑 **Export to PDF** *(Optional)*  
  After inserting data, you can export the same file as PDF with the same filename.

- ⚡ **Reusable & Lightweight**  
  Designed for enterprise workflows with high reliability.

---

## 🛠 Activities

### `InsertDataToBookmarks`
Inserts a list of key-value pairs into predefined bookmarks in a Word document.

### `ConvertWordToPdf`
Converts a Word document into a PDF, preserving formatting and structure.

### `FillBookmarksAndExport`
Single-step activity that fills bookmarks and exports to PDF in one go.

---

## 🧾 Inputs

| Name             | Type         | Description                                                              |
|------------------|--------------|--------------------------------------------------------------------------|
| FilePath         | String       | Path to the input Word document.            |
| Bookmarks        | Dictionary<String,String> | A list of bookmarks and their corresponding values.                        |
| SaveAsPath       | String       | Custom output path for the updated Word document.                         |
| ConvertToPdf     | Boolean      | If `True`, a PDF version will also be created with the same filename.     |

---

## 📤 Outputs

- **output** *(String)* –  
  - If `SaveAsPath` is provided → returns that path.  
  - If `SaveAsPath` is **not provided** → the input file is overwritten and the original `FilePath` is returned.
  - In case of error, the error will be returned.   

---

## ⚙️ Dependencies

- .NET 6 Runtime
- UiPath Studio 2022.10+

---

## 💡 Use Cases

- Business analysts or users can ask questions about databases without writing SQL.
- Generate reports or summaries from live data with Gemini.
- Automate insight generation from structured data.

---

## 🔧 How to Install

1. Download the `.nupkg` file from the [Releases](https://github.com/AmrEzzatAbdo/Word.BookmarkWriter.Activities/releases/).
2. Add the local path in UiPath Studio: `Manage Packages > Settings > User-defined package sources`.
3. Go to `Manage Packages > Available > Your Feed`, and install the activity.

---

## 🖥 Usage Example

1. Prepare a Word template with bookmarks (e.g., `CustomerName`, `AccountNumber`).  
2. Create a list or dictionary of key-value pairs in UiPath:  

```vb
New Dictionary(Of String, String) From
{
    {"CustomerName", "John Doe"},
    {"AccountNumber", "1234567890"}
}
```

3. Use the `InsertDataToBokmarks` activity to inject values and export to PDF.

---


## 📜 License

This project is licensed under the [MIT License].

---

## 👨‍💻 Author

Developed by **Amr Ezzat**

- LinkedIn: [linkedin.com/in/amr-ezzat](https://www.linkedin.com/in/amrezzatabdal-al/)
- Email: `amrezzat289@gmail.com`

---

> Feel free to contribute or open issues for improvements!
