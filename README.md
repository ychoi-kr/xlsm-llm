# xlsm-llm

Excel VBA functions utilizing local LLMs.

This repository provides a set of Excel VBA modules that integrate with local language model (LLM) servers. It allows you to send prompts and receive processed text (e.g., summaries, code generation, text correction) directly into Excel, complete with error handling and proper formatting.

## Features

- **LLM**: Send a prompt to the LLM server and receive a response.
- **LLM_SUMMARIZE**: Summarize a given text using the LLM.
- **LLM_CODE**: Generate code based on provided requirements.
- **LLM_LIST**: Create a list with formatted output.
- **LLM_EDIT**: Correct and improve a given sentence for clarity, grammar, and punctuation.
- Modularized helper functions:
  - **BuildJsonPayload**: Constructs the JSON payload for the API request.
  - **SendLLMRequest**: Sends HTTP requests with detailed error reporting.
  - **ExtractContent**: Uses regex to extract the response content without trailing noise.
- Automatic newline conversion for proper cell formatting.
- **API Key Support**: Each function accepts an optional API key (as the last parameter) that lets you authenticate with external LLM APIs (such as OpenAI, Gemini, Upstage, etc.) by including the appropriate token.

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/ychoi-kr/xlsm-llm.git
   ```
2. Import the VBA modules from the `src` directory (e.g., `LLM_Functions.bas`) into your Excel workbook.

## Usage

- Use `=LLM(prompt, [value], [temperature], [max_tokens], [model], [base_url], [show_think], [api_key])` to get responses from your local LLM server.
    ![](img/usage_LLM.png)
- Use `=LLM_SUMMARIZE(text, [prompt], [temperature], [max_tokens], [model], [base_url], [show_think], [api_key])` to generate summaries.
    ![](img/usage_LLM_SUMMARIZE.png)
- Use `=LLM_CODE(program_detail, programming_language, [model], [base_url], [show_think], [api_key])` to write code.
    ![](img/usage_LLM_CODE.png)
- Use `=LLM_LIST(prompt, [model], [base_url], [show_think], [api_key])` to create a list.
    ![](img/usage_LLM_LIST.png)
- Use `=LLM_EDIT(text, [prompt], [temperature], [max_tokens], [model], [base_url], [show_think], [api_key])` to correct and edit sentences.  
  By default, the function uses the prompt:  
  `Please correct the following sentence for clarity, grammar, and punctuation:`  
  You can override this by providing a custom prompt if desired.  
    ![](img/usage_LLM_EDIT.png)

- Ensure your server URL is correctly configured, or pass it as the optional `base_url` parameter.

## License

This project is licensed under the MIT License.
