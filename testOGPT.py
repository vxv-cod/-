# import os
# import openai


# openai.organization = "https://api.openai.com/v1/models/Home"
# openai.api_key = os.getenv("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")
# openai.Model.list()



import os
import openai

# openai.api_key = "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"
openai.api_key = os.getenv("OPENAI_API_KEY")

response = openai.Completion.create(
    model="text-davinci-003",
    # model="text-davinci-002",
    # model="text-davinci-001",
    # model="text-davinci-001",
    # prompt="Россия",
    prompt="что скажешь о холодильниках Hair?",
    temperature=0.9,
    max_tokens=1000,
    top_p=1,
    frequency_penalty=0.0,
    presence_penalty=0.6,
    stop=[" Human:", " AI:"]
)

print(response["choices"][0]["text"])


# проверь код