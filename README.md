# 👟 AI-Powered Spec Sheet Generator

An intelligent system that automatically generates footwear specification sheets from development sample logs using advanced AI. Built with OpenAI GPT-4.1 for smart material matching and complex text parsing.

## 🚀 Quick Start

### Option 1: One-Click Setup & Launch
```bash
python3 setup_and_run.py
```

### Option 2: Manual Setup
```bash
# Install dependencies
pip3 install -r requirements_full.txt

# Launch with ngrok tunnel
python3 run_with_ngrok.py

# Or run locally
streamlit run app.py
```

## 🎯 Features

### 🤖 AI-Enhanced Processing
- **Smart Material Matching**: GPT-4.1 understands synonyms, abbreviations, and variations
- **Complex Text Parsing**: Handles messy supplier descriptions like "W1063 Minnetonka Brown Cow Suede - lining: Microfiber"
- **Multi-layer Fallback**: Exact → Fuzzy → AI matching for maximum accuracy
- **Confidence Scoring**: Shows how certain each match is

### 📊 Robust Processing
- **Batch Processing**: Handle hundreds of samples at once
- **Error Recovery**: Continues processing even if individual samples fail
- **Detailed Statistics**: Track exact, fuzzy, and AI matches
- **Processing Logs**: Full visibility into what's happening

### 🌐 Beautiful Web Interface
- **Drag & Drop Upload**: Easy file management
- **Live Preview**: See your data before processing
- **Progress Tracking**: Real-time processing updates
- **One-Click Download**: Get your results instantly

## 📋 File Requirements

### 1. Development Sample Log (Excel)
- **Format**: First row = headers, subsequent rows = data
- **Required columns**: 
  - Sample Name
  - Season  
  - Gender
  - Factory ref # (or Factory Ref #)
  - Sample Order No.
  - Upper (material descriptions)
  - Sole (ref # only) or Sole

### 2. Spec Template (Excel)
- Your blank specification sheet template
- The system will duplicate this for each sample

### 3. Simplified BOM (Excel)
- Your standardized material vocabulary
- Format: "Part: Material Name" or just material names

## 🔧 How It Works

### Processing Pipeline
```
📄 Raw Input → 🤖 AI Parsing → 🎯 Smart Matching → 📝 Spec Generation → 📥 Download
```

### AI Enhancement Levels

1. **Exact Match** (100% confidence)
   - Perfect string matches with BOM

2. **Fuzzy Match** (85% confidence)  
   - Similar strings with 85%+ similarity

3. **AI Match** (70-95% confidence)
   - GPT-4.1 analyzes context and meaning
   - Handles synonyms, abbreviations, typos
   - Understands material relationships

### Example AI Processing

**Input**: `"Upper: W1063 Minnetonka Brown Cow Suede - lining: Microfiber, Color: Brown"`

**AI Parsing**:
```json
{
  "Upper": "Cow Suede",
  "Lining": "Microfiber"
}
```

**AI Matching**: 
- "Cow Suede" → "Suede Leather" (AI: 85% confidence)
- "Microfiber" → "Microfiber" (Exact: 100% confidence)

## 🛠️ Technical Architecture

### Backend (`ai_spec_processor.py`)
- **AISpecProcessor**: Main processing engine
- **MaterialMatch**: Structured match results with confidence
- **ProcessingResult**: Comprehensive processing statistics
- **Caching**: Avoid repeated AI calls for same materials

### Frontend (`app.py`)
- **Streamlit**: Modern web interface
- **File Upload**: Multi-file drag & drop
- **Progress Tracking**: Real-time updates
- **Results Dashboard**: Statistics & download

### Models Used
- **GPT-4.1**: Complex material matching and reasoning
- **GPT-4.1-mini**: Text parsing and extraction (cost-optimized)

## 💡 AI vs Traditional Comparison

| Feature | Traditional (Regex) | AI-Enhanced |
|---------|-------------------|-------------|
| **Exact matches** | ✅ Perfect | ✅ Perfect |
| **Typos** | ❌ Fails | ✅ Handles |
| **Abbreviations** | ❌ Limited | ✅ Understands |
| **Synonyms** | ❌ No context | ✅ Contextual |
| **Complex descriptions** | ❌ Rigid patterns | ✅ Flexible parsing |
| **Multi-language** | ❌ English only | ✅ Multi-lingual |
| **Cost** | ✅ Free | 💰 API costs |
| **Speed** | ✅ Instant | ⏱️ ~1-2s per sample |

## 🔐 Security & Privacy

- **API Key**: Stored locally, never transmitted except to OpenAI
- **File Processing**: All processing happens locally
- **No Data Storage**: Files are processed in memory only
- **Secure Transmission**: HTTPS for all API calls

## 🚀 Deployment Options

### Local Development
```bash
streamlit run app.py
```

### ngrok Tunnel (Testing)
```bash
python3 run_with_ngrok.py
```

### Production Deployment
- **Streamlit Cloud**: Easy deployment with GitHub integration
- **Heroku**: Container-based deployment
- **AWS/Azure**: Full cloud deployment
- **Docker**: Containerized deployment

## 📊 Performance & Costs

### Processing Speed
- **Exact matches**: Instant
- **Fuzzy matches**: ~0.1s per material
- **AI matches**: ~1-2s per material (API latency)
- **Overall**: ~10-30 seconds for 50 samples

### OpenAI API Costs (Approximate)
- **GPT-4.1-mini**: ~$0.001 per material match
- **Typical job** (50 samples, 200 materials): ~$0.20
- **Large job** (500 samples, 2000 materials): ~$2.00

## 🛠️ Customization

### Adding New Material Types
Edit `ai_spec_processor.py`:
```python
patterns = [
    r'(Upper[^:]*?):\s*([^-\n]+?)(?:\s*[-\n]|$)',
    r'(YourNewPart[^:]*?):\s*([^,\n]+?)(?:\s*[,\n]|$)',  # Add here
]
```

### Adjusting AI Confidence
In the web interface or code:
```python
confidence_threshold = 0.7  # Adjust as needed
```

### Custom Prompts
Modify prompts in `_ai_material_match()` and `parse_material_block()` methods.

## 🐛 Troubleshooting

### Common Issues

**"ModuleNotFoundError: No module named 'openai'"**
```bash
pip3 install openai>=1.12.0
```

**"Invalid API Key"**
- Check your OpenAI API key is correct
- Ensure you have sufficient credits
- Verify API key has appropriate permissions

**"No materials found in BOM"**
- Check BOM file format
- Ensure materials are in the expected columns
- Try the "Part: Material" format

**ngrok Issues**
```bash
# Install ngrok separately if needed
brew install ngrok  # macOS
# or download from https://ngrok.com/download
```

### Debug Mode
Set logging level for detailed output:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Add tests if applicable  
5. Submit a pull request

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- **OpenAI** for GPT-4.1 and excellent API
- **Streamlit** for the amazing web framework
- **ngrok** for easy tunneling solutions

---

**Made with ❤️ for footwear professionals who deserve better tools** 