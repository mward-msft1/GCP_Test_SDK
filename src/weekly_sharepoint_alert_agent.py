from __future__ import annotations

import base64
import json
import os
from dataclasses import dataclass
from typing import Any, Dict, Optional
from urllib.parse import quote

try:
    import requests
    from azure.identity import DefaultAzureCredential
    from dotenv import load_dotenv
    from flask import Flask, jsonify
except ImportError:  # pragma: no cover - dependency installation is environment-specific.
    requests = None
    DefaultAzureCredential = None
    load_dotenv = lambda: None

    class Flask:
        def __init__(self, *args, **kwargs):
            pass

        def get(self, *args, **kwargs):
            def decorator(func):
                return func
            return decorator

        def post(self, *args, **kwargs):
            def decorator(func):
                return func
            return decorator

        def run(self, *args, **kwargs):
            raise RuntimeError("Install the project requirements before starting the Flask app.")

    def jsonify(payload):
        return payload

    Flask = Flask

load_dotenv()

app = Flask(__name__)


@dataclass
class AgentSettings:
    app_name: str = "weekly-sharepoint-compliance-agent"
    tenant_id: str = ""
    client_id: str = ""
    client_secret: str = ""
    sharepoint_site_id: str = ""
    sharepoint_site_url: str = ""
    sharepoint_file_path: str = ""
    teams_team_id: str = ""
    teams_channel_id: str = ""
    email_to: str = ""
    email_from: str = ""
    purview_client_app_id: str = ""
    purview_app_name: str = "My Secure Agent"
    azure_openai_endpoint: str = ""
    azure_openai_chat_completion_model: str = "gpt-4o-mini"


def load_settings() -> AgentSettings:
    return AgentSettings(
        app_name=os.getenv("APP_NAME", "weekly-sharepoint-compliance-agent"),
        tenant_id=os.getenv("TENANT_ID", ""),
        client_id=os.getenv("CLIENT_ID", ""),
        client_secret=os.getenv("CLIENT_SECRET", ""),
        sharepoint_site_id=os.getenv("SHAREPOINT_SITE_ID", ""),
        sharepoint_site_url=os.getenv("SHAREPOINT_SITE_URL", "https://contoso.sharepoint.com/sites/LegalOps"),
        sharepoint_file_path=os.getenv("SHAREPOINT_FILE_PATH", "Shared Documents/Compliance/Quarterly-Review.docx"),
        teams_team_id=os.getenv("TEAMS_TEAM_ID", ""),
        teams_channel_id=os.getenv("TEAMS_CHANNEL_ID", ""),
        email_to=os.getenv("EMAIL_TO", "alerts@contoso.com"),
        email_from=os.getenv("EMAIL_FROM", "agent@contoso.com"),
        purview_client_app_id=os.getenv("PURVIEW_CLIENT_APP_ID", ""),
        purview_app_name=os.getenv("PURVIEW_APP_NAME", "My Secure Agent"),
        azure_openai_endpoint=os.getenv("AZURE_OPENAI_ENDPOINT", ""),
        azure_openai_chat_completion_model=os.getenv("AZURE_OPENAI_CHAT_COMPLETION_MODEL", "gpt-4o-mini"),
    )


class MicrosoftGraphClient:
    def __init__(self, settings: AgentSettings):
        if requests is None or DefaultAzureCredential is None:
            raise RuntimeError("Install the project requirements with 'pip install -r requirements.txt' before running the workflow.")
        self.settings = settings
        self.credential = DefaultAzureCredential()

    def _get_access_token(self) -> str:
        token = self.credential.get_token("https://graph.microsoft.com/.default")
        return token.token

    def _request(self, method: str, path: str, json_body: Optional[Dict[str, Any]] = None, params: Optional[Dict[str, Any]] = None, headers: Optional[Dict[str, str]] = None, timeout: int = 60):
        url = f"https://graph.microsoft.com/v1.0{path}"
        request_headers = {"Authorization": f"Bearer {self._get_access_token()}", "Accept": "application/json"}
        if headers:
            request_headers.update(headers)
        response = requests.request(method=method.upper(), url=url, headers=request_headers, json=json_body, params=params, timeout=timeout)
        if response.status_code >= 400:
            raise RuntimeError(f"Graph API error for {method} {path}: {response.status_code} {response.text}")
        if not response.content:
            return {}
        try:
            return response.json()
        except ValueError:
            return {"raw_text": response.text}

    def build_document_url(self) -> str:
        encoded_path = quote(self.settings.sharepoint_file_path, safe="/")
        base_url = self.settings.sharepoint_site_url.rstrip("/")
        file_path = self.settings.sharepoint_file_path.strip("/")
        return f"{base_url}/{file_path}"

    def get_document_bytes(self) -> Dict[str, Any]:
        encoded_path = quote(self.settings.sharepoint_file_path, safe="/")
        path = f"/sites/{self.settings.sharepoint_site_id}/drive/root:/{encoded_path}:/content"
        response = requests.get(
            f"https://graph.microsoft.com/v1.0{path}",
            headers={"Authorization": f"Bearer {self._get_access_token()}"},
            timeout=60,
        )
        if response.status_code >= 400:
            raise RuntimeError(f"Unable to download document from SharePoint: {response.status_code} {response.text}")
        return {
            "file_name": os.path.basename(self.settings.sharepoint_file_path),
            "content": response.content,
            "content_type": response.headers.get("Content-Type", "application/octet-stream"),
            "document_url": self.build_document_url(),
        }

    def send_teams_message_with_document(self, document_url: str, file_name: str) -> Dict[str, Any]:
        body = {
            "body": {
                "contentType": "html",
                "content": (
                    "<p>Weekly review alert: a new SharePoint document is ready for compliance review.</p>"
                    f"<p><a href='{document_url}'>{file_name}</a></p>"
                ),
            },
            "attachments": [
                {
                    "@odata.type": "#microsoft.graph.chatMessageAttachment",
                    "contentType": "reference",
                    "contentUrl": document_url,
                    "name": file_name,
                }
            ],
        }
        return self._request("POST", f"/teams/{self.settings.teams_team_id}/channels/{self.settings.teams_channel_id}/messages", json_body=body)

    def send_email_with_document(self, file_name: str, file_bytes: bytes) -> Dict[str, Any]:
        message = {
            "message": {
                "subject": "Weekly SharePoint compliance alert",
                "body": {
                    "contentType": "Text",
                    "content": "The attached document requires review. It was retrieved by the weekly SharePoint compliance agent.",
                },
                "toRecipients": [{"emailAddress": {"address": self.settings.email_to}}],
                "from": {"emailAddress": {"address": self.settings.email_from}},
                "attachments": [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": file_name,
                        "contentType": "application/octet-stream",
                        "contentBytes": base64.b64encode(file_bytes).decode("utf-8"),
                    }
                ],
            },
            "saveToSentItems": "true",
        }
        return self._request("POST", f"/users/{self.settings.email_from}/sendMail", json_body=message)

    def get_security_alerts(self) -> Dict[str, Any]:
        return self._request("GET", "/security/alerts?$top=10")


class PurviewDlpMonitor:
    def __init__(self, graph_client: MicrosoftGraphClient):
        self.graph = graph_client

    def evaluate_document(self, file_name: str, content_bytes: bytes) -> Dict[str, Any]:
        safe_content = content_bytes.decode("utf-8", errors="ignore")
        findings = []
        sensitive_markers = [
            "credit card",
            "ssn",
            "password",
            "secret",
            "confidential",
        ]
        for marker in sensitive_markers:
            if marker.lower() in safe_content.lower():
                findings.append(marker)

        alerts = self.graph.get_security_alerts()
        return {
            "file_name": file_name,
            "sensitive_markers_found": findings,
            "dlp_alert_summary": alerts.get("value", [])[:5],
            "status": "review-required" if findings else "no-manual-review-required",
        }


def build_agent_framework_agent() -> Optional[Any]:
    try:
        from agent_framework import Agent
        from agent_framework.microsoft import PurviewPolicyMiddleware, PurviewSettings
        from azure.identity import InteractiveBrowserCredential
    except Exception:
        return None

    settings = load_settings()
    if not settings.purview_client_app_id:
        return None

    credential = InteractiveBrowserCredential(client_id=settings.purview_client_app_id)
    middleware = [PurviewPolicyMiddleware(credential=credential, settings=PurviewSettings(app_name=settings.purview_app_name))]
    # This is a scaffold used to follow the Agent Framework pattern; the actual runtime client is user-defined.
    # Replace with your preferred chat client if you want to wire in live Azure AI inference.
    return {
        "name": settings.app_name,
        "middleware": middleware,
        "instructions": "Review the SharePoint document, block risky content according to Purview policy, and notify security stakeholders when required.",
    }


def run_weekly_workflow() -> Dict[str, Any]:
    settings = load_settings()
    try:
        graph_client = MicrosoftGraphClient(settings)
        document = graph_client.get_document_bytes()
        purview = PurviewDlpMonitor(graph_client)
        review = purview.evaluate_document(document["file_name"], document["content"])

        teams_result = graph_client.send_teams_message_with_document(document["document_url"], document["file_name"])
        email_result = graph_client.send_email_with_document(document["file_name"], document["content"])

        result = {
            "app_name": settings.app_name,
            "document_name": document["file_name"],
            "document_url": document["document_url"],
            "sensitive_markers_found": review["sensitive_markers_found"],
            "dlp_alert_summary": review["dlp_alert_summary"],
            "teams_message": teams_result,
            "email_result": email_result,
            "agent_framework_ready": build_agent_framework_agent() is not None,
        }
        return result
    except RuntimeError as exc:
        return {"status": "not-ready", "error": str(exc)}


@app.get("/health")
def health_check():
    return jsonify({"status": "ok", "service": "weekly-sharepoint-alert-agent"})


@app.post("/run")
def trigger_run():
    try:
        payload = run_weekly_workflow()
        return jsonify({"status": "success", "result": payload}), 200
    except Exception as exc:  # pragma: no cover - central error handling for Cloud Run
        return jsonify({"status": "error", "message": str(exc)}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT", "8080")))
