import runpy
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import MagicMock, patch
import pytest
import pandas
import bs4
import openpyxl

def load():
    openai = MagicMock()
    server = MagicMock()
    server.search.return_value = ('OK', [b'1'])
    server.fetch.return_value = ('OK', [(b'1', b'Subject: Ordinary message\nDate: Tue, 06 Oct 2026 12:00:00 +0000\n\nNo alert')])
    dependencies = {name:MagicMock() for name in ['pydrive','pydrive.auth','pydrive.drive']}
    dependencies['openai'] = openai
    with patch.dict('sys.modules', dependencies), patch('imaplib.IMAP4_SSL', return_value=server):
        module = runpy.run_path(str(Path(__file__).resolve().parents[1] / 'Githubgptv1.4.py'))
    return module, openai

@pytest.mark.parametrize('answer,expected', [(' YES \n',True),('no',False),('',False),('maybe',False)])
def test_article_classification_normalizes_response(answer, expected):
    module, openai = load()
    openai.Completion.create.return_value = SimpleNamespace(choices=[SimpleNamespace(text=answer)])
    assert module['is_article']('https://example.test/article') is expected
    args = openai.Completion.create.call_args.kwargs
    assert 'https://example.test/article' in args['prompt']
    assert args['max_tokens'] == 1

def test_provider_error_is_not_reported_as_article():
    module, openai = load()
    openai.Completion.create.side_effect = RuntimeError('offline')
    with pytest.raises(RuntimeError, match='offline'):
        module['is_article']('https://example.test')
