from textblob import TextBlob
import openai

# Initialize OpenAI API
openai.api_key = ""

def correct_text(text):
    # Correct orthographic errors using TextBlob
    blob = TextBlob(text)
    corrected_text = str(blob.correct())
    return corrected_text

def change_tone(text):
    # Modify the tone using OpenAI's GPT-4 model
    response = openai.ChatCompletion.create(
        model="gpt-4",  # Use "gpt-4" or "gpt-3.5-turbo" based on your access
        messages=[
            {"role": "system", "content": "You are a helpful assistant."},
            {"role": "user", "content": f"Rewrite the following text in a casual yet professional tone: \"{text}\""}
        ],
        temperature=0.7
    )
    return response.choices[0].message['content'].strip()

# Example function that combines both
def process_text(input_text):
    corrected = correct_text(input_text)
    professional_tone = change_tone(corrected)
    return professional_tone

# Test the function
input_description = input('Enter a description:')
result = process_text(input_description)
print(result)
