# xlsm-llm

Excel VBA functions for interacting with local and cloud-based LLMs, enabling text processing, translation, summarization, and code generation directly in Excel.  
These functions allow you to send prompts to LLMs and retrieve AI-generated responses directly in your spreadsheet, supporting both local models and external APIs like OpenAI, Gemini, and Upstage.

## Features  

xlsm-llm provides powerful Excel VBA functions that seamlessly integrate with local and external LLMs, enabling text processing, translation, summarization, and code generation directly within Excel.

**✨ Interact with LLMs directly in Excel**
- **LLM:** Send a prompt and receive AI-generated responses in your spreadsheet.
- **LLM_SUMMARIZE:** Automatically generate concise summaries of long texts.
- **LLM_CODE:** Generate code snippets in any programming language based on your requirements.
- **LLM_LIST:** Extract structured lists from AI responses for better data organization.
- **LLM_EDIT:** Improve text by fixing grammar, punctuation, and clarity.
- **LLM_TRANSLATE:** Translate text into different languages, with custom translation prompts available.

**⚡ Optimized for Ease of Use**
- Works with both **local LLMs and cloud APIs** (e.g., OpenAI, Gemini, Upstage).
- API keys can be provided **manually** or set as **environment variables** for seamless authentication.
- Functions handle **newline formatting** automatically, ensuring clean Excel outputs.

**✅ Reliable and Flexible**
- Supports **custom prompts** for greater control over AI-generated outputs.
- Includes **error handling** for stable and predictable function execution.
- Compatible with **various LLM models**, including GPT-based models, Upstage Solar, and locally hosted models.

With xlsm-llm, you can harness AI capabilities without leaving Excel, enhancing productivity and automation in text processing and coding tasks.

## Installation

1. **Clone the repository:**
   ```sh
   git clone https://github.com/ychoi-kr/xlsm-llm.git
   ```
2. **Import the VBA modules into your Excel workbook:**
   - Open your Excel workbook.
   - Press `Alt + F11` to open the Visual Basic for Applications (VBA) editor.
   - In the VBA editor, press `Ctrl + R` if the Project Explorer is not already visible. Locate your workbook's project (e.g., `VBAProject (YourWorkbookName.xlsm)`).
   - Right-click on your workbook's project in the Project Explorer and select **Import File…**.
   - Navigate to the `src` folder in the cloned repository, select the module file (for example, `LLM_Functions.bas`), and click **Open**.
   - The module should now appear in your Project Explorer.
3. **Add the required Microsoft XML reference:**
   - In the VBA editor, go to **Tools > References**.
   - Check **Microsoft XML, v6.0** from the list and click **OK**.

## Setting Environment Variables

Before using the functions, ensure that you have obtained an API key from your desired service provider (e.g., OpenAI, Gemini, Upstage, etc.). These keys allow you to authenticate with external LLM APIs. You can store these keys as environment variables so that the functions automatically use them when the API key parameter is omitted.

For example, if you are using the OpenAI API, set the environment variable `OPENAI_API_KEY`. Similarly, for Gemini or Upstage, you can set `GEMINI_API_KEY` or `UPSTAGE_API_KEY`, respectively.

- **On Windows:**  
  - Open **Control Panel > System > Advanced system settings > Environment Variables**.  
  - Under "User variables" (or "System variables" for all users), click **New…** and set the variable name (e.g., `OPENAI_API_KEY`) and the value as your API key.  
  - Alternatively, use the command prompt:  
    ```sh
    setx OPENAI_API_KEY your_api_key_here
    ```
- **On macOS/Linux:**  
  - Edit your shell profile (e.g., `~/.bash_profile` or `~/.zshrc`) and add a line such as:  
    ```sh
    export OPENAI_API_KEY=your_api_key_here
    ```  
  - Then reload your profile by running `source ~/.bash_profile` (or the appropriate command for your shell).

This method works for any API key required by the functions; just replace `OPENAI_API_KEY` with the appropriate variable name for the service you are using (e.g., `GEMINI_API_KEY`, `UPSTAGE_API_KEY`).

## Using a Local LLM

You can use `xlsm-llm` with a **local LLM** instead of a cloud-based API (e.g., OpenAI, Gemini, Upstage). To do this, you need to **set up and run a local LLM server**.  

The easiest way to run a local LLM is with **[LM Studio](https://lmstudio.ai/)**:  

1. **Download and install LM Studio** from [lmstudio.ai](https://lmstudio.ai/).  
2. **Download a model** – LM Studio suggests recommended models upon first launch, or you can browse and install a model of your choice.  
3. **Start the local server** in the **Local Server** tab.  
4. **Set the `base_url` in Excel functions** to match the server’s address (e.g., `http://localhost:1234/v1/`).  
5. **Specify the model explicitly** when calling the function, as local LLMs require a model name.  

For example:  
```
=LLM("Write a haiku", "recursion in programming", 0.7, 100, "qwen2.5-14b-instruct", "http://localhost:1234/v1/")
```

💡 **No API key is required** for most local LLMs.  

For detailed setup instructions, refer to the [LM Studio documentation](https://lmstudio.ai/docs).

## Usage

### LLM
This function sends a prompt to your local LLM server and returns the processed response. You can either let the function use the API key stored in your environment variable or provide one directly as the last parameter.

**Example format:**
```
=LLM(prompt, [value], [temperature], [max_tokens], [model], [base_url], [show_think], [api_key])
```
If you omit the `api_key` argument, the function will automatically use the API key stored in the relevant environment variable (e.g., `OPENAI_API_KEY` for OpenAI API).

![](img/usage_LLM.png)

> **Note:** If the response includes line breaks, you may need to enable text wrapping in the cell. To do this, right-click the cell, select **Format Cells**, go to the **Alignment** tab, and check **Wrap Text**.

### LLM_SUMMARIZE
LLM_SUMMARIZE processes the input text to produce a succinct summary. It is ideal for quickly condensing lengthy texts.

**Example format:**
```
=LLM_SUMMARIZE(text, [prompt], [temperature], [max_tokens], [model], [base_url], [show_think], [api_key])
```
![](img/usage_LLM_SUMMARIZE.png)

### LLM_CODE
LLM_CODE generates code based on the provided program details and programming language. This function is perfect for automating code snippet creation.

**Example format:**
```
=LLM_CODE(program_detail, programming_language, [model], [base_url], [show_think], [api_key])
```
![](img/usage_LLM_CODE.png)

### LLM_LIST
LLM_LIST creates a list of items using a specified prompt. The output is formatted using special tags, ensuring consistency in the list's presentation.

**Example format:**
```
=LLM_LIST(prompt, [model], [base_url], [show_think], [api_key])
```
![](img/usage_LLM_LIST.png)

### LLM_EDIT
LLM_EDIT refines a sentence by correcting grammar, punctuation, and overall clarity. By default, it uses the prompt:  
`Please correct the following sentence for clarity, grammar, and punctuation:`  
However, you can supply a custom prompt to tailor the editing process to your needs.

**Example format:**
```
=LLM_EDIT(text, [prompt], [temperature], [max_tokens], [model], [base_url], [show_think], [api_key])
```
![](img/usage_LLM_EDIT.png)

### LLM_TRANSLATE
LLM_TRANSLATE translates text from one language to another using your LLM. The simplest usage requires only the text and the target language. You can also provide a custom translation instruction if desired.

- **General Usage:**  
  In the simplest case, you specify the text to be translated and the target language. Optionally, you can include the source language and/or a custom prompt. When a custom prompt is provided, both the `targetLang` and `sourceLang` parameters are ignored.
  
  **Format:**
  ```
  =LLM_TRANSLATE(text, targetLang, [sourceLang], [customPrompt], [temperature], [maxTokens], [model], [base_url], [show_think], [api_key])
  ```
  
  For example, to translate "Hello, world!" into Spanish:
  ```
  =LLM_TRANSLATE("Hello, world!", "es")
  ```
  
  And to use a custom translation instruction (which ignores `targetLang` and `sourceLang`):
  ```
  =LLM_TRANSLATE(text, [targetLang ignored], [sourceLang ignored], customPrompt, [temperature], [maxTokens], [model], [base_url], [show_think], [api_key])
  ```

- **Upstage API Exception:**  
  When using Upstage API's `translation-enko` or `translation-koen` models, the function ignores the `targetLang`, `sourceLang`, and `customPrompt` parameters, using only the input text for translation.
  
  **Format:**
  ```
  =LLM_TRANSLATE(text, [targetLang ignored], [sourceLang ignored], [customPrompt ignored], [temperature], [maxTokens], [model], [base_url], [show_think], [api_key])
  ```
  
![](img/usage_LLM_TRANSLATE.png)

## License

This project is licensed under the MIT License.
