import openai


def send_prompt(key: str, prompt: str) -> str | None:
    openai.api_key = key

    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {"role": "user", "content": prompt}
        ],
        max_tokens=100,
        temperature=0.7
    )

    return response["choices"][0]["message"]["content"]
