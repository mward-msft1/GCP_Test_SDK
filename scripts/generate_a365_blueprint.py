#!/usr/bin/env python3
"""Generate a simple agent blueprint for the weekly SharePoint alert workflow."""

from __future__ import annotations

import json
from pathlib import Path

BLUEPRINT = {
    "name": "weekly-sharepoint-compliance-agent",
    "description": "Monitors a SharePoint document, checks for sensitive information indicators, shares it in Teams, emails it, and logs DLP alerts.",
    "runtime": {
        "platform": "gcp-cloud-run",
        "language": "python",
        "schedule": "cron: 0 9 * * 1"
    },
    "auth": {
        "mode": "service-principal",
        "provider": "microsoft-entra-id",
        "appRegistration": {
            "displayName": "weekly-sharepoint-compliance-agent",
            "requiredGraphPermissions": [
                "Files.Read.All",
                "Sites.Read.All",
                "Mail.Send",
                "ChannelMessage.Send",
                "SecurityEvents.Read.All",
                "TeamworkAppInstallation.ReadWriteForTeam.All"
            ],
            "adminConsentRequired": True
        }
    },
    "steps": [
        "Resolve the target SharePoint document.",
        "Download and inspect the file contents.",
        "Run Purview DLP checks and log sensitive data findings.",
        "Send a Teams message with the document URL or attachment.",
        "Send an email with the file attached.",
        "Write a summary to app logs and optional monitoring alerts."
    ],
    "variables": {
        "APP_NAME": "weekly-sharepoint-compliance-agent",
        "TENANT_ID": "<tenant-id>",
        "CLIENT_ID": "<client-id>",
        "CLIENT_SECRET": "<client-secret>",
        "SHAREPOINT_SITE_ID": "<site-id>",
        "SHAREPOINT_FILE_PATH": "Shared Documents/Compliance/Quarterly-Review.docx",
        "EMAIL_TO": "alerts@contoso.com",
        "TEAMS_TEAM_ID": "<team-id>",
        "TEAMS_CHANNEL_ID": "<channel-id>",
        "PURVIEW_CLIENT_APP_ID": "<purview-client-id>"
    }
}


def main() -> None:
    output_path = Path(__file__).resolve().parent.parent / "generated" / "agent_blueprint.json"
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(BLUEPRINT, indent=2), encoding="utf-8")
    print(f"Blueprint written to {output_path}")


if __name__ == "__main__":
    main()
