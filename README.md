Here is a full **README file** for your final Python script. It includes:

* ✅ Project overview
* 🔧 Requirements
* 📥 Sample input Excel
* 📤 Sample output Excel
* 💡 Usage instructions
* 📁 Sample files description (with tables)

---

### 📄 `README.md`

````markdown
# 🧠 ConversationID Bot Tracker

This Python script fetches conversation details using an API and generates a clean Excel report with:

- Last **User** message before an empty **Bot** response.
- Last **Bot** message before the empty **Bot**.
- Most recent **URL** seen before the empty **Bot** response.

The output Excel is formatted with:
- Proper column widths
- Blue-colored headers

---

## 📦 Requirements

Install dependencies before running:

```bash
pip install pandas openpyxl requests
````

---

## 📥 Sample Input File (`conversation_ids.xlsx`)

| cr7de\_conversationid   |
| ----------------------- |
| 12345678-abcd-4321-efgh |
| 98765432-zyxw-8765-ijkl |

✅ Column name **must** be: `cr7de_conversationid`

---

## ⚙️ How It Works

1. Loads conversation IDs from Excel.
2. Calls the API to retrieve full conversation details.
3. Locates the first empty Bot message.
4. Extracts the most recent User message, Bot message, and URL before that.
5. Saves the final output in a clean, readable Excel file.

---

## 📤 Sample Output File (`conversation_ids_output.xlsx`)

| ConversationId          | User                            | Bot                            | currentUrlLink                                             |
| ----------------------- | ------------------------------- | ------------------------------ | ---------------------------------------------------------- |
| 12345678-abcd-4321-efgh | “What are your business hours?” | “Our hours are 9am to 5pm.”    | [https://example.com/contact](https://example.com/contact) |
| 98765432-zyxw-8765-ijkl | “Can you show latest pricing?”  | “Sure, here’s our price list.” | [https://example.com/pricing](https://example.com/pricing) |

🟦 **Header row is colored blue**
📏 **All columns are sized to fit on one PC screen**

---

## 🚀 Running the Script

```bash
python bot.py
```

Make sure your input file is named: `conversation_ids.xlsx`

---

## 📁 File Overview

| File Name                      | Description                              |
| ------------------------------ | ---------------------------------------- |
| `bot_conversation_tracker.py`  | Main script to fetch and generate Excel. |
| `conversation_ids.xlsx`        | Input file with `cr7de_conversationid`.  |
| `conversation_ids_output.xlsx` | Output report with formatted data.       |

---

## 📌 Notes

* Script will skip any conversation where no empty Bot response is found.
* If an API error occurs for a row, the script logs the issue and continues.

---

## 📧 Contact

For improvements or questions, feel free to reach out.

```

---

### 📎 Sample Excel Files (Preview)

#### ✅ Sample Input (`conversation_ids.xlsx`)

| cr7de_conversationid     |
|--------------------------|
| `1111-aaaa-2222-bbbb`    |
| `3333-cccc-4444-dddd`    |

#### ✅ Sample Output (`conversation_ids_output.xlsx`)

| ConversationId         | User                     | Bot                    | currentUrlLink            |
|------------------------|--------------------------|-------------------------|---------------------------|
| `1111-aaaa-2222-bbbb`  | How do I reset password? | Click "Forgot Password" | https://example.com/help  |
| `3333-cccc-4444-dddd`  | Where's my invoice?      | Sent to your email      | https://example.com/invoice |

*All headers in light blue, and column widths fit nicely on screen.*

---

Would you like me to generate these sample Excel files for download as well?
```
