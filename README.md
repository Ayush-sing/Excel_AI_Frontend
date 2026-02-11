# 📊 Excel AI Assistant

An intelligent Excel Add-in that transforms natural language commands into real-time data analysis, visualizations, and regression outputs directly inside Microsoft Excel.

---

## 🚀 Overview

Excel AI Assistant allows users to interact with spreadsheet data using simple, structured natural language commands. Instead of writing formulas or manually creating charts, users can request insights such as:

- Aggregations (sum, mean, median, etc.)
- Charts (bar, line, scatter, box, histogram, pie, heatmap)
- Regression analysis (linear, polynomial, Logarithmic)
- Filtered analysis using `where` conditions

All results are generated dynamically and placed back into Excel.

---

## 🧠 How It Works

The system combines:

- **Intent Classification (Machine Learning)** for detecting user requests
- **Rule-Based Column Resolution** using `{column_name}` syntax
- **FastAPI Backend** for processing
- **React + TypeScript Excel Add-in Frontend**
- **Pandas, NumPy & Matplotlib** for analytics and visualization

This hybrid design ensures:
- Accuracy
- Reduced ambiguity
- Structured and reliable Excel operations

---

## ✨ Key Features

- 📈 Create charts using natural language  
- 📊 Perform statistical calculations instantly  
- 🔎 Apply filters using `where` conditions  
- 📉 Run regression analysis with visual output  
- 📁 Upload Excel/CSV files for dynamic processing  
- 🧾 Generate formatted results directly in Excel  

---

## 📝 Example Commands

```bash
sum {total_profit}

sum {total_profit} where {shopping_mode} is Online

mean {sales} where {region} is West

scatter plot of {unit_cost} and {unit_price}

box plot of {profit} by {category}

regression of {ad_spend} and {sales}
```
Column Name must be specified between {}
---

## 📂 Documentation & Examples

Detailed usage instructions and sample commands are available here:

📄 [Excel AI Assistant Examples & Manual](docs/assets/Excel_AI_Assistant_Examples.pdf)

---

## How you can use it?

Download manifest.excel.xml and use instructions shown in the video below
Source: Youtube | Youtuber: Michael Zlatkovsky (Not me)

[![Watch the video](https://img.youtube.com/vi/XXsAw2UUiQo/maxresdefault.jpg)](https://youtu.be/XXsAw2UUiQo)

### [Watch this video on YouTube](https://youtu.be/XXsAw2UUiQo)

## 🛠 Deployment Architecture

- Backend deployed using Render (FastAPI + Uvicorn)
- Frontend hosted via GitHub Pages
- Excel Add-in distributed through sideloaded `manifest.xml`

---

## 🎯 Purpose

This project demonstrates:

- End-to-end AI-powered application development
- Integration of machine learning with business tools
- Production-ready API deployment
- Excel automation through Office Add-ins

---

## 📌 Use Cases

- Business data analysis
- Academic data exploration
- Sales performance tracking
- Operational reporting
- Quick statistical insights without formulas

---

## ⚠️ Limitations

- Requires exact column names inside `{}`  
- Not designed for advanced forecasting models
- Cannot work well with dates
- Does not perform fuzzy column matching  
- Focused on structured numerical datasets  

---

## 📬 Contact

For collaboration, demo access, or technical discussion, feel free to connect.

---

**Excel AI Assistant — Bringing AI-powered analytics directly into Excel.**
