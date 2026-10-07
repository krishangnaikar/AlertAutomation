## Tests

```sh
python -m pip install -r requirements-test.txt
python -m pytest
```

These initial tests cover selected behavior with external services mocked. They do not establish full integration coverage.

Coverage: article-classification response normalization and API error propagation. IMAP is mocked during module loading. Email parsing, spreadsheets, and live API integration are not covered.
