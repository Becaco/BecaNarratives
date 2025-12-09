# BecaNarratives 🎨✨

## ✨ Features

- 🤖 **5 AI Models**: Groq, OpenAI, Gemini, Claude, Phi-4
- 🗄️ **Neo4j Integration**: Direct extraction from knowledge graphs
- 💾 **Smart Saving**: Auto-saves every 10 objects
- 🔄 **Resume Support**: Continue from where you left off
- 📊 **Excel Output**: Beautiful formatted spreadsheets
- ⚡ **Fast & Reliable**: Automatic retry and rate limit handling

---

## 📦 Installation

```bash
# Basic installation
pip install pandas openpyxl groq openai langchain-openai google-generativeai anthropic

# Add Neo4j support (optional)
pip install langchain-community py2neo
```

---

## 🎯 Quick Start

### 1️⃣ Set your API key

```bash
export TOKEN="your_api_key_here"
```

### 2️⃣ Extract data from Neo4j (optional)

```bash
python extract_data.py
```

### 3️⃣ Generate stories

```bash
python story_generator.py
```
---

## 💡 Usage Examples


```python
from story_generator import process_catalog

process_catalog(
    input_file="catalog.xlsx",
    output_file="stories.xlsx",
    provider='groq'
)
```

---

## 🗄️ Neo4j Extraction

Connect to your knowledge graph and extract object data:

```python
from story_generator import Neo4jExtractor

extractor = Neo4jExtractor(
    url="bolt://localhost:7687",
    username="neo4j",
    password="your_password"
)

# Extract 100 objects
extractor.extract_to_excel(
    output_file="catalog.xlsx",
    limit=100
)
```

Or simply use the script:

```bash
python extract_data.py
```

---


## 📝 Input/Output Format

### Input Excel

| N° | Object ID | Description |
|----|-----------|-------------|
| 1 | OBJ_001 | Bronze vessel from 500 BCE... |
| 2 | OBJ_002 | Ceramic bowl with geometric patterns... |

### Output Excel

Same as input + **Narrative Story** column (170-270 words of engaging narrative)

---

## 🎨 Story Structure

Each generated story follows this proven structure:

1. **Opening Hook** (25-35 words) - Grab attention
2. **Physical Description** (20-30 words) - Help visualize
3. **Historical Context** (50-70 words) - Inform
4. **Historical Journey** (60-80 words) - Tell the story
5. **Present Connection** (25-30 words) - Relate to today

---

## ⚙️ Configuration

```python
process_catalog(
    input_file="catalog.xlsx",     # Your Excel file
    output_file="stories.xlsx",    # Output with stories
    provider='groq',               # AI provider
    model=None,                    # Override default model
    api_key=None,                  # Or use TOKEN env var
    save_interval=10               # Save every N objects
)
```

---

## 🔑 Get API Keys

- **Groq**: https://console.groq.com/
- **OpenAI**: https://platform.openai.com/
- **Gemini**: https://ai.google.dev/
- **Claude**: https://console.anthropic.com/
- **Phi-4**: https://github.com/marketplace/models

 

 
 
 

