# ExcelFlow - AI-Powered Excel Assistant

<div align="center">

![ExcelFlow Logo](https://img.shields.io/badge/ExcelFlow-AI%20Excel%20Assistant-blue?style=for-the-badge)

[![FastAPI](https://img.shields.io/badge/FastAPI-005571?style=for-the-badge&logo=fastapi)](https://fastapi.tiangolo.com/)
[![React](https://img.shields.io/badge/React-61DAFB?style=for-the-badge&logo=react&logoColor=black)](https://reactjs.org/)
[![OpenAI](https://img.shields.io/badge/OpenAI-412991?style=for-the-badge&logo=openai&logoColor=white)](https://openai.com/)
[![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/)
[![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)](https://developer.mozilla.org/en-US/docs/Web/JavaScript)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

</div>

An interactive web application for managing and analyzing Excel sheets using OpenAI's GPT models. This app allows users to ask natural language questions about their Excel file, perform updates, summarize data, and more — all with the help of a conversational AI.

**If you find this project useful, please consider giving it a star ⭐ on GitHub!**

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

* **Real-time Updates:**
  See changes to your Excel file in real-time as the AI makes modifications.

* **Modern Web Interface:**
  React-based frontend with a responsive design for a great user experience.

* **WebSocket Communication:**
  Real-time bidirectional communication between the frontend and backend.

---

## 🏗️ Architecture

The application is split into two main components:

### Backend (FastAPI)
- RESTful API for file uploads and data retrieval
- WebSocket server for real-time communication
- Excel manipulation utilities
- OpenAI integration for natural language processing

### Frontend (React)
- Modern React application built with Vite
- Interactive spreadsheet view using React Data Grid
- Real-time chat interface
- File upload functionality

---

## 🛠️ Setup

### Backend Setup

1. **Navigate to the backend directory:**

   ```bash
   cd backend
   ```

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Create and fill your `.env` file in the backend directory:**

   ```ini
   OPENAI_API_KEY=your_openai_key
   ```

4. **Run the backend server:**

   ```bash
   uvicorn app.main:app --reload
   ```

### Frontend Setup

1. **Navigate to the frontend directory:**

   ```bash
   cd frontend
   ```

2. **Install dependencies:**

   ```bash
   npm install
   ```

3. **Run the development server:**

   ```bash
   npm run dev
   ```

4. **Open your browser and navigate to:**
   
   ```
   http://localhost:3000
   ```

## 📝 Usage

1. Upload an Excel file through the web interface
2. View your spreadsheet in the main area
3. Use the chat interface on the right to interact with the AI
4. Ask questions or give commands about your data
5. Watch as changes are reflected in real-time in the spreadsheet view

### Example Commands

- "What's the sum of sales in column C?"
- "Find all cells containing 'Product X' and replace with 'Product Y'"
- "Insert a new row at position 5"
- "Calculate the average of cells B5:B15"
- "Add a total row at the bottom of the sheet"
- "Sort column D in descending order"
- "Highlight all cells with values greater than 100"

## 📊 Demo

Coming Soon!

## 🔧 Advanced Configuration

### Environment Variables

**Backend (.env file):**

```ini
OPENAI_API_KEY=your_openai_key
DEBUG=True  # Optional for verbose logging
```

### Customizing the AI Model

You can change the OpenAI model used by modifying the `model` parameter in `backend/app/agent.py`:

```python
response = self.client.chat.completions.create(
    model="gpt-4o",  # Change to your preferred model
    messages=all_messages,
    tools=self.tools,
    tool_choice="auto"
)
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 👨‍💻 Author

**Parth Vadhadiya**

- Website: [https://parthvadhadiya.netlify.app/](https://parthvadhadiya.netlify.app/)
- LinkedIn: [https://www.linkedin.com/in/parth-vadhadiya/](https://www.linkedin.com/in/parth-vadhadiya/)
- GitHub: [@parthvadhadiya](https://github.com/parthvadhadiya)

---

<div align="center">

### If you found this project helpful, please consider giving it a star! ⭐

</div>
