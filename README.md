# PowerPoint Translator

PowerPoint Translator is a Python3 script that uses Azure OpenAI's GPT-4o language model to translate the text in a PowerPoint presentation from English to Traditional Chinese.

## Getting Started

### Prerequisites

To run this script, you'll need the following:

- Python3 (version 3.11 or higher)
- `openai` Python package (version 3.0.0 or higher)
- `pptx` Python package (version 1.0.2 or higher)
- `streamlit` Python package (version 1.20.0 or higher)

- A valid Azure OpenAI Service API key from [Microsoft Azure Subscription](https://azure.microsoft.com/en-us/products/ai-services/openai-service)

### Installing

1. Clone or download the PowerPoint Translator repository to your local machine.
2. Install the required Python packages by running `pip install -r requirements.txt` in your terminal or command prompt.

### Setting up the environment variables for Azure OpenAI Service
This script uses the Python dotenv package. Environment variables can also be written in the .env file, for example:

```bash
OPENAI_API_KEY=<your Azure OpenAI API key ex:1234567890abcdef1234567890abcdef>
OPENAI_API_BASE=https://<your Azure OpenAI resource name>.openai.azure.com/
DEPLOYMENT_NAME=gpt-4o
OPENAI_API_VERSION=2023-03-15-preview
OPENAI_API_TYPE=azure
```
To put the .env file with main.py in the same folder.

### Usage

1. Open a terminal or command prompt and navigate to the directory where you saved the PowerPoint Translator repository.
2. Run the command `streamlit run main.py` to start the script.
3. Upload the PowerPoint file you want to translate.
4. Wait for the script to translate the text in your PowerPoint file. A progress bar will show you the progress of the translation process.
5. When the script has finished, click the "Download"
