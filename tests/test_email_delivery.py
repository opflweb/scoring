from pathlib import Path

from opfl.constants import ALL_TEAM_CODES
from scripts.email_delivery import (
    TEAM_EMAIL_VARS,
    all_recipients,
    commissioner_recipients,
    recipients_for_teams,
    send_email,
)

PROJECT_ROOT = Path(__file__).resolve().parent.parent


def test_every_team_has_a_secret_safe_email_var():
    assert set(TEAM_EMAIL_VARS) == set(ALL_TEAM_CODES)
    assert TEAM_EMAIL_VARS['K/D'] == 'K_D_EMAIL'
    for var in TEAM_EMAIL_VARS.values():
        assert var.replace('_', '').isalnum()


def test_coowner_addresses_split_and_dedupe(monkeypatch):
    monkeypatch.setenv('K_D_EMAIL', 'kirk@example.com, david@example.com')
    monkeypatch.setenv('G_G_EMAIL', 'kirk@example.com')

    assert recipients_for_teams(['K/D']) == ['david@example.com', 'kirk@example.com']
    assert recipients_for_teams(['K/D', 'G/G', 'NOPE']) == ['david@example.com', 'kirk@example.com']
    assert set(all_recipients()) >= {'david@example.com', 'kirk@example.com'}


def test_disabled_email_mode_routes_everything_to_commissioner(monkeypatch):
    monkeypatch.setenv('DISABLE_EMAILS', 'true')
    monkeypatch.setenv('COMMISSIONER_EMAIL', 'commissioner@example.com')
    monkeypatch.setenv('K_D_EMAIL', 'manager@example.com')

    assert recipients_for_teams(['K/D']) == ['commissioner@example.com']
    assert all_recipients() == ['commissioner@example.com']
    assert commissioner_recipients() == ['commissioner@example.com']


def test_send_email_without_credentials_fails_closed(monkeypatch):
    monkeypatch.delenv('SMTP_USERNAME', raising=False)
    monkeypatch.delenv('SMTP_PASSWORD', raising=False)

    assert send_email('subject', 'body', ['a@example.com']) is False
    assert send_email('subject', 'body', []) is False


def test_score_workflow_alerts_commissioner_on_failure():
    source = (PROJECT_ROOT / '.github' / 'workflows' / 'score.yml').read_text(encoding='utf-8')

    assert 'if: failure()' in source
    assert 'from scripts.email_delivery import commissioner_recipients, send_email' in source
    assert 'SMTP_USERNAME: ${{ secrets.SMTP_USERNAME }}' in source
    assert 'SMTP_PASSWORD: ${{ secrets.SMTP_PASSWORD }}' in source
