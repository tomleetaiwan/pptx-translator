# PowerPoint Translator

PowerPoint Translator is a Streamlit application that uses the official OpenAI Python SDK 1.x to translate text in a PowerPoint presentation into Traditional Chinese. This version connects to Azure OpenAI by using Microsoft Entra ID authentication.

This codebase is currently tested against an Azure OpenAI Service deployment based on the `gpt-5-mini` model.

## Getting Started

### Prerequisites

To run this script, you'll need the following:

- Python 3.11 or higher
- `openai` Python package (version 1.109.1 or higher)
- `azure-identity` Python package (version 1.17.1 or higher)
- `python-pptx` Python package (version 1.0.2 or higher)
- `streamlit` Python package (version 1.39.0 or higher)

- An Azure OpenAI resource with a deployed chat model
- A Microsoft Entra identity that can access that Azure OpenAI resource

### Installing

1. Clone or download the PowerPoint Translator repository to your local machine.
2. Install the required Python packages by running `python -m pip install -r requirements.txt`.

If you prefer a global Python environment, install the dependencies into that exact interpreter. For example:

```bash
python.exe -m pip install -r requirements.txt
```

### Setting up authentication for Microsoft Entra ID
This project uses `python-dotenv`, so you can place the required environment variables in a `.env` file in the same folder as `main.py`.

1. Sign in with an identity that can access the Azure OpenAI resource.

For local development, Azure CLI is the simplest option:

```bash
az login
```

2. Make sure the signed-in identity has the `Cognitive Services OpenAI User` role on the target Azure OpenAI resource.

3. Configure the application with the Azure OpenAI endpoint and deployment name.

Example:

```bash
OPENAI_MODEL=<your deployment name>
OPENAI_BASE_URL=https://<your-resource-name>.openai.azure.com/openai/v1/
```

For example, if you created an Azure OpenAI deployment for `gpt-5-mini`, set `OPENAI_MODEL` to that deployment name. It only needs to be the literal string `gpt-5-mini` if that is the deployment name you chose.

You can also keep using:

```bash
OPENAI_ENDPOINT=https://<your-resource-name>.openai.azure.com/
```

When `OPENAI_ENDPOINT` is set, the app automatically converts it to the Azure OpenAI v1 `base_url` format.

Notes:

- `OPENAI_MODEL` is required. For Azure OpenAI, this value must be your deployment name.
- `OPENAI_BASE_URL` is required unless you provide `OPENAI_ENDPOINT` instead.
- `OPENAI_API_VERSION` is not required for Azure OpenAI v1 and is ignored by this app.
- `OPENAI_API_KEY` is not used in this version because authentication is handled by Microsoft Entra ID.
- `DefaultAzureCredential` can use Azure CLI sign-in, environment-based service principal credentials, managed identity, shared token cache, or interactive browser sign-in.

### Usage

1. Open a terminal and change to the project directory.
2. Start the app with the same Python interpreter that has the dependencies installed.

For a global Python environment:

```bash
python.exe -m streamlit run main.py
```

Do not rely on `streamlit run main.py` by itself unless you are sure that `streamlit` on your PATH belongs to the same Python installation.

3. Upload the PowerPoint file you want to translate.
4. Optionally fill in `Optional slide context and translation notes` to describe the deck's topic, terminology, or translation constraints. The app adds this text to the system prompt for every translation request.
5. Optionally fill in `Optional terminology glossary` with one term per line using `source => target`. Matching source terms are forced to the target translation across the deck.
6. Review `Suggested glossary candidates` to see repeated English terms extracted from the uploaded deck. Select any terms you want to preserve first, then click `Add selected terms to glossary` to append `source => source` draft entries that you can edit.
7. Click `Check Entra sign-in` before translation. If this machine doesn't already have a cached Microsoft Entra sign-in, the app can open your default browser so you can sign in first.
8. Review the `Microsoft Entra Status` panel. It shows whether sign-in is ready, whether a browser prompt is expected, whether translation is currently enabled, and what to do next.
9. Click `Translate PowerPoint` after you finish reviewing the glossary and the Entra sign-in check succeeds. The app will show a progress bar while it processes the presentation.
10. Click `Download` to save the translated PowerPoint file.

Notes:

- If the same source string appears multiple times in the same upload, the app translates that exact string once and reuses the result across the deck. This improves consistency for repeated sentences and reduces duplicate model calls.
- Glossary entries are applied after the model returns, so terms such as product names or domain-specific phrases keep the exact target translation you specify.
- Suggested glossary candidates are heuristic English term matches based on repeated text in the deck. Review them before adding them to the glossary.
- `Check Entra sign-in` verifies that the app can acquire a Microsoft Entra ID token before translation starts. Azure OpenAI RBAC permissions are still enforced when the translation request is sent.
