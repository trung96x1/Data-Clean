from deep_translator import GoogleTranslator

text = "СТЕАРИНОВАЯ КИСЛОТА, ЕЕ СОЛИ И СЛОЖНЫЕ ЭФИPЫ:"

# Automatically detects language and translates to English
translated = GoogleTranslator(source='auto', target='en').translate(text)

print(translated)  # Output: Hello world