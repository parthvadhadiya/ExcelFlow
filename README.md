# Excel Agent with OpenAI

An interactive, Streamlit-powered assistant for managing and analyzing Excel sheets using OpenAI's GPT models. This app allows users to ask natural language questions about their Excel file, perform updates, summarize data, and more — all with the help of a conversational AI.

---

## 🚀 Features

* **Chat with your Excel Sheet:**
  Ask questions, filter data, make edits, and get summaries in plain English.

* **Read & Update Cells/Ranges:**
  Read individual cells or ranges, update data, insert/delete rows and columns, and more.

* **Summarization:**
  Quickly get sums, averages, mins, and maxes for any range of numbers.

* **Find & Replace:**
  Perform powerful find-and-replace operations over the whole sheet or just a selected area.

* **OpenAI GPT-4o Integration:**
  Advanced reasoning and understanding through OpenAI's GPT models and function calling.

* **Secure Environment Variables:**
  Configuration (like API keys and Excel path) via `.env` file.

* **No-Code Data Operations:**
  All common data operations accessible through a chat-based UI — no formulas needed!

---

## 🛠️ Setup

1. **Clone the repository:**

   ```bash
   git clone <your-repo-url>
   cd <your-repo-folder>
   ```

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Create and fill your `.env` file:**

   ```ini
   OPENAI_API_KEY=your_openai_key
   EXCEL_PATH=./data.xlsx   # Path to your Excel file
   ```

4. **Run the app:**

   ```bash
   streamlit run <your-script>.py
   ```
