# doc2zipf

See if your Word document conforms to Zipf's Law!

## Requirements

* Python 3
* [python-docx](https://python-docx.readthedocs.io/en/latest/)

## Usage

```
python doc2zipf.py [DOCUMENT...]

Outputs a Word document containing a table with a list of words and the number
of times they occur as "sorted - [DOCUMENT]"
eg. "sorted - Literature Review.docx"

DOCUMENT:
    A valid Word document in .docx format (Word 2007 onwards)
	You can specify multiple documents if you want.
```

## Disclaimer

This may or may not be 100% accurate as I have only tested it on a
couple of documents myself. Should be fine for general use, as it was
only really made out of curiosity.