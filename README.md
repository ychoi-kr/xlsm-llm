# xlsm-llm

Excel VBA functions utilizing local LLMs.

This repository provides a set of Excel VBA modules that integrate with local language model (LLM) servers. It allows you to send prompts and receive processed text (e.g., summaries) directly into Excel, complete with error handling and proper formatting.

## Features

- **LLM**: Send a prompt to the LLM server and receive a response.
- **LLM_SUMMARIZE**: Summarize a given text using the LLM.
- Modularized helper functions:
  - **BuildJsonPayload**: Constructs the JSON payload for the API request.
  - **SendLLMRequest**: Sends HTTP requests with detailed error reporting.
  - **ExtractContent**: Uses regex to extract the response content without trailing noise.
- Automatic newline conversion for proper cell formatting.

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/your_username/xlsm-llm.git
   ```
2. Import the VBA modules from the `src` directory (e.g., `LLM_Functions.bas`) into your Excel workbook.

## Usage

- Use `=LLM(prompt, [value], [temperature], [max_tokens], [model], [base_url])` to get responses from your local LLM server.
- Use `=LLM_SUMMARIZE(text, [prompt], [temperature], [max_tokens], [model], [base_url])` to generate summaries.
- Ensure your server URL is correctly configured, or pass it as the optional `base_url` parameter.

## License

This project is licensed under the MIT License.
