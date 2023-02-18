"""
pip install transformers
"""
from transformers import pipeline

summarizers = {
    # 'es': "philschmid/bart-large-cnn-samsum",
    # 'en': "facebook/bart-large-cnn"
    'en': "philschmid/bart-large-cnn-samsum",
}


def load_pipeline(language):
    language = language.lower()
    assert language in summarizers, f"""Language {language} is not currently implemented: try modifying summarizer.py. Available languages: {list(summarizers.keys())}"""
    checkpoint = summarizers['en']
    summarizer = pipeline("summarization", model=checkpoint)
    return summarizer


def summarize(text, summarizer, max_length=None):
    if max_length:
        return summarizer(text, max_length=max_length)
    else:
        return summarizer(text)


if __name__ == "__main__":
    text = """The Sun is the star at the center of the Solar System. It is a nearly perfect ball of hot plasma, heated to incandescence by nuclear fusion reactions in its core. The Sun radiates this energy mainly as light, ultraviolet, and infrared radiation, and is the most important source of energy for life on Earth."""
    summarizer = load_pipeline('en')
    result = summarize(text, summarizer, max_length=min(120, int(len(compiled_text) * 0.3)))
    print(result)
