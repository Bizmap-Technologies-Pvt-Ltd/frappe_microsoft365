"""A fake Microsoft Graph at the one network seam, for tests that must run everything above it.

Patching ``graph_request`` hides graph_url, the URL guard, the 401/429 handling and the error
text every message in this app is built from, and 1.1.1 shipped a regression through exactly
that gap. These helpers patch only ``microsoft_graph._http_request`` and the token lookups, so a
test sees the method, URL and body that would really go out.
"""

import json
from contextlib import contextmanager
from unittest.mock import patch

from frappe_microsoft365 import microsoft_graph as graph

REASONS = {200: "OK", 403: "Forbidden", 404: "Not Found", 429: "Too Many Requests", 503: "Service Unavailable"}


class Response:
	"""The parts of a requests.Response that graph_request reads."""

	def __init__(self, status=200, body=None, content=None, headers=None):
		self.status_code = status
		self.headers = headers or {}
		self.reason = REASONS.get(status, "")
		if content is None:
			content = json.dumps(body).encode() if body is not None else b""
		self.content = content

	def json(self):
		return json.loads(self.content or b"{}")

	def close(self):
		pass

	def iter_content(self, chunk_size=None):
		return iter([self.content])


def error(status, code, message=""):
	return Response(status, {"error": {"code": code, "message": message or code}})


@contextmanager
def fake_graph(router):
	"""Answer every Graph call with ``router(method, url, kwargs)``; yields the list of calls.

	Each call is recorded as ``(method, url, kwargs)`` exactly as graph_request passed it on.
	"""
	calls = []

	def http_request(method, url, **kwargs):
		calls.append((method, url, kwargs))
		return router(method, url, kwargs)

	with patch.object(graph, "_http_request", side_effect=http_request), patch.object(
		graph, "get_valid_access_token", return_value="token"
	), patch.object(graph, "get_app_access_token", return_value="app-token"), patch(
		"time.sleep"
	):
		yield calls
