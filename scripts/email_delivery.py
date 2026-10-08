#!/usr/bin/env python3
"""Shared SMTP delivery and recipient lookup for OPFL emails.

Ported from qpfl/scoring's scripts/email_delivery.py. Sends through the same
Gmail bot account (SMTP_USERNAME / SMTP_PASSWORD secrets).
"""

import os
import smtplib
from email.message import EmailMessage

from opfl.constants import ALL_TEAM_CODES


def team_email_var(team: str) -> str:
    """Secret name holding a team's address(es): 'K/D' -> 'K_D_EMAIL'."""
    return f'{team.replace("/", "_")}_EMAIL'


TEAM_EMAIL_VARS = {team: team_email_var(team) for team in ALL_TEAM_CODES}

# Alerts and DISABLE_EMAILS test mode go here instead of to the league.
COMMISSIONER_EMAIL_VAR = 'COMMISSIONER_EMAIL'


def emails_disabled() -> bool:
    return os.environ.get('DISABLE_EMAILS', '').lower() == 'true'


def _addresses(*email_vars: str) -> set[str]:
    # Each secret may hold a comma-separated list, for co-owned teams.
    return {
        address.strip()
        for email_var in email_vars
        for address in os.environ.get(email_var, '').split(',')
        if address.strip()
    }


def commissioner_recipients() -> list[str]:
    return sorted(_addresses(COMMISSIONER_EMAIL_VAR))


def recipients_for_teams(teams: list[str]) -> list[str]:
    """Return de-duplicated addresses for one or more league teams."""
    if emails_disabled():
        return commissioner_recipients()
    return sorted(_addresses(*(TEAM_EMAIL_VARS[team] for team in teams if team in TEAM_EMAIL_VARS)))


def all_recipients() -> list[str]:
    """Return every team's addresses, or the commissioner when emails are disabled."""
    return recipients_for_teams(ALL_TEAM_CODES)


def send_email(subject: str, body: str, recipients: list[str]) -> bool:
    if not recipients:
        print(f'No email address configured for {subject}')
        return False

    smtp_user = os.environ.get('SMTP_USERNAME')
    smtp_password = os.environ.get('SMTP_PASSWORD')
    if not smtp_user or not smtp_password:
        print('SMTP_USERNAME or SMTP_PASSWORD is not configured')
        return False

    message = EmailMessage()
    message['Subject'] = subject
    message['From'] = f'OPFL Bot <{smtp_user}>'
    message['To'] = ', '.join(recipients)
    message.set_content(body)

    try:
        with smtplib.SMTP('smtp.gmail.com', 587, timeout=30) as server:
            server.starttls()
            server.login(smtp_user, smtp_password)
            server.send_message(message)
        return True
    except Exception as error:
        print(f'Could not send {subject}: {error}')
        return False
