"""Resolve a File's owner to the canonical SharePoint storage anchor.

Consumer apps register one provider per source DocType in ``hooks.py``::

    microsoft365_storage_target_resolver = {
        "Work Note": "my_app.storage.resolve_work_note_target",
    }

The provider receives ``(source_doctype, source_name)`` and returns ``None`` to opt out,
or ``{"anchor_doctype": ..., "anchor_name": ..., "path_segments": [...]} ``.
Omit ``path_segments`` to use the anchor mapping's normal base-folder and folder-pattern
rules. Providers must be deterministic, read-only, and return the final existing anchor.
This module never accepts a client-selected target or Graph ID.
"""

from __future__ import annotations

import json
import threading
from dataclasses import dataclass
from typing import Any

import frappe
from frappe import _

BINDING_FIELD = "custom_microsoft_storage_binding"
BINDING_VERSION = 1
MAX_FOLDER_NAME = 120
_RESOLVING = threading.local()


class StorageTargetError(frappe.ValidationError):
	"""A configured storage target is missing, ambiguous or no longer matches its binding."""


@dataclass(frozen=True)
class StorageTarget:
	source_doctype: str
	source_name: str
	anchor_doctype: str
	anchor_name: str
	mapping: Any
	path_segments: tuple[str, ...] | None
	dynamic: bool
	folder: Any = None

	@property
	def path_mode(self):
		return "explicit" if self.path_segments is not None else "default"


def _provider_map():
	return frappe.get_hooks("microsoft365_storage_target_resolver", {}) or {}


def source_doctypes():
	"""DocTypes registered for dynamic storage; used by retry and the Files panel boot."""
	return sorted(_provider_map())


def _normalize_path(segments):
	from frappe_microsoft365.microsoft_files import safe_name
	if not isinstance(segments, list) or not segments:
		frappe.throw(_("A custom storage path must be a non-empty list of folder names."), StorageTargetError)
	out = []
	for segment in segments:
		if not isinstance(segment, str) or not segment or segment != segment.strip():
			frappe.throw(_("Storage path segments must be non-empty trimmed strings."), StorageTargetError)
		if segment in (".", "..") or "/" in segment or "\\" in segment:
			frappe.throw(_("A storage path segment cannot contain a path or relative component."), StorageTargetError)
		clean = safe_name(segment, fallback="", max_length=MAX_FOLDER_NAME)
		if not clean:
			frappe.throw(
				_("Storage path segment {0} cannot be converted to a SharePoint folder name.").format(segment),
				StorageTargetError,
			)
		out.append(clean)
	if sum(len(part) for part in out) + len(out) - 1 > 400:
		frappe.throw(_("The custom SharePoint path exceeds 400 characters."), StorageTargetError)
	return tuple(out)


def _providers_for(doctype):
	value = _provider_map().get(doctype, [])
	if isinstance(value, str):
		return [value]
	return list(value or [])


def _call_provider(doctype, name, provider):
	stack = getattr(_RESOLVING, "stack", ())
	key = (doctype, name)
	if key in stack:
		frappe.throw(_("A storage-target resolver recursively resolved {0} {1}.").format(doctype, name), StorageTargetError)
	_RESOLVING.stack = (*stack, key)
	try:
		try:
			return frappe.get_attr(provider)(doctype, name)
		except StorageTargetError:
			raise
		except Exception as e:
			frappe.throw(
				_("Storage-target provider for {0} failed: {1}").format(doctype, str(e)),
				StorageTargetError,
			)
	finally:
		_RESOLVING.stack = stack


def _provider_result(doctype, name):
	providers = _providers_for(doctype)
	if len(providers) > 1:
		frappe.throw(
			_("More than one Microsoft 365 storage-target provider is registered for {0}.").format(doctype),
			StorageTargetError,
		)
	if not providers:
		return False, None
	return True, _call_provider(doctype, name, providers[0])


def resolve_storage_target(source_doctype, source_name):
	"""Return the current canonical target, or ``None`` when this file should stay local.

	With no provider, the existing mapping for the source remains the exact default. A provider
	that returns ``None`` explicitly opts this source record out; a configured provider error
	or invalid result fails closed.
	"""
	from frappe_microsoft365 import microsoft_files as files

	dynamic, result = _provider_result(source_doctype, source_name)
	if dynamic:
		if result is None:
			return None
		if not isinstance(result, dict) or set(result) - {"anchor_doctype", "anchor_name", "path_segments"}:
			frappe.throw(_("The storage-target provider returned an invalid response."), StorageTargetError)
		anchor_doctype = result.get("anchor_doctype")
		anchor_name = result.get("anchor_name")
		if not isinstance(anchor_doctype, str) or not anchor_doctype or not isinstance(anchor_name, str) or not anchor_name:
			frappe.throw(_("The storage-target provider must return an anchor DocType and name."), StorageTargetError)
		if not frappe.db.exists(anchor_doctype, anchor_name):
			frappe.throw(_("Storage anchor {0} {1} does not exist.").format(anchor_doctype, anchor_name), StorageTargetError)
		if (anchor_doctype, anchor_name) == (source_doctype, source_name):
			frappe.throw(
				_("A storage-target provider must return a distinct final anchor; use the default mapping for the source itself."),
				StorageTargetError,
			)
		# An anchor is final. Reject provider chaining so the mapping and permissions have one
		# unambiguous owner and a chain cannot recurse through two or more business apps.
		chained, next_result = _provider_result(anchor_doctype, anchor_name)
		if chained and next_result is not None:
			frappe.throw(
				_("Storage anchor {0} {1} is itself redirected by a provider; return the final anchor.").format(
					anchor_doctype, anchor_name
				),
				StorageTargetError,
			)
		path = _normalize_path(result["path_segments"]) if result.get("path_segments") is not None else None
	else:
		anchor_doctype, anchor_name = source_doctype, source_name
		path = None

	mapping = files.mapping_for(anchor_doctype)
	if not mapping:
		if dynamic:
			frappe.throw(
				_("No enabled SharePoint Mapping exists for storage anchor DocType {0}.").format(anchor_doctype),
				StorageTargetError,
			)
		return None
	if path is not None:
		all_segments = [*files.base_segments(mapping), *path]
		if sum(len(part) for part in all_segments) + max(0, len(all_segments) - 1) > 400:
			frappe.throw(_("The SharePoint path including the mapping base folder exceeds 400 characters."), StorageTargetError)
	return StorageTarget(
		source_doctype=source_doctype,
		source_name=source_name,
		anchor_doctype=anchor_doctype,
		anchor_name=anchor_name,
		mapping=mapping,
		path_segments=path,
		dynamic=dynamic,
	)


def binding_for(target, folder=None):
	"""Canonical v1 snapshot. Source remains File.attached_to_* and is not duplicated."""
	value = {
		"v": BINDING_VERSION,
		"anchor": [target.anchor_doctype, target.anchor_name],
		"path": {"mode": target.path_mode},
		"folder": None,
	}
	if target.path_segments is not None:
		value["path"]["segments"] = list(target.path_segments)
	if folder:
		value["folder"] = {"drive_id": folder.drive_id, "item_id": folder.item_id}
	return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def parse_binding(raw):
	if not raw:
		return None
	try:
		value = json.loads(raw)
		if (
			set(value) != {"v", "anchor", "path", "folder"}
			or type(value["v"]) is not int
			or value["v"] != BINDING_VERSION
			or not isinstance(value["anchor"], list)
			or len(value["anchor"]) != 2
			or not all(isinstance(v, str) and v for v in value["anchor"])
			or not isinstance(value["path"], dict)
			or value["path"].get("mode") not in ("default", "explicit")
		):
			raise ValueError
		if value["path"]["mode"] == "default":
			if set(value["path"]) != {"mode"}:
				raise ValueError
		else:
			if set(value["path"]) != {"mode", "segments"}:
				raise ValueError
			value["path"]["segments"] = list(_normalize_path(value["path"]["segments"]))
		if value["folder"] is not None and (
			not isinstance(value["folder"], dict)
			or set(value["folder"]) != {"drive_id", "item_id"}
			or not all(isinstance(v, str) and v for v in value["folder"].values())
		):
			raise ValueError
		return value
	except Exception:
		frappe.throw(_("The File has an invalid Microsoft 365 storage binding; access is blocked."), StorageTargetError)


def assert_binding_matches(file_doc, target, require_dynamic_proof=True):
	"""Fail closed if a resolver/File binding is missing, corrupt or logically different."""
	if not target:
		if file_doc.get(BINDING_FIELD):
			frappe.throw(_("The File's storage resolver no longer returns a target; access is blocked."), StorageTargetError)
		return None
	binding = parse_binding(file_doc.get(BINDING_FIELD))
	if binding is None:
		if target.dynamic and require_dynamic_proof:
			frappe.throw(
				_("This legacy File has no storage-binding history for the current dynamic target; reconciliation is required."),
				StorageTargetError,
			)
		return None
	path = binding["path"]
	expected = {"mode": target.path_mode}
	if target.path_segments is not None:
		expected["segments"] = list(target.path_segments)
	if binding["anchor"] != [target.anchor_doctype, target.anchor_name] or path != expected:
		frappe.throw(
			_("The File's stored SharePoint target no longer matches the resolver; automatic moves are blocked and reconciliation is required."),
			StorageTargetError,
		)
	return binding


def authorize_target(target, permission, user=None):
	"""Require the same Frappe permission on the source and its storage anchor."""
	user = user or frappe.session.user
	frappe.has_permission(target.source_doctype, permission, doc=target.source_name, user=user, throw=True)
	frappe.has_permission(target.anchor_doctype, permission, doc=target.anchor_name, user=user, throw=True)


def validate_file_context(file_doc, permission=None, require_folder=True):
	"""Re-resolve and validate stored logical and physical identity for a File access.

	An unbound legacy file keeps the old default path. It cannot cross into a dynamic target
	without upload-time history. Bound files require both current logical equality and the
	anchor's still-persisted folder ID to match the snapshot.
	"""
	from frappe_microsoft365 import microsoft_files as files

	source_doctype = file_doc.get("attached_to_doctype")
	source_name = file_doc.get("attached_to_name")
	current = resolve_storage_target(source_doctype, source_name)
	stored = assert_binding_matches(file_doc, current)
	if not current:
		return None
	# Unbound Files retain the historical File permission behavior. The additional Source and
	# Anchor checks apply to resolver-bound Files, where File's URL can otherwise outlive context.
	if permission and stored:
		authorize_target(current, permission)
	if not stored:
		return current
	if not current.dynamic:
		frappe.throw(_("The File's storage resolver was removed; its stored target is no longer authorized."), StorageTargetError)
	physical = stored.get("folder")
	if not physical:
		if require_folder:
			frappe.throw(_("The File's storage binding has no resolved folder; access is blocked."), StorageTargetError)
		return current
	folder = files.get_folder(current.anchor_doctype, current.anchor_name)
	if not folder or (folder.drive_id, folder.item_id) != (physical["drive_id"], physical["item_id"]):
		frappe.throw(
			_("The File's original SharePoint folder binding is missing or changed; reconciliation is required."),
			StorageTargetError,
		)
	return StorageTarget(
		**{**current.__dict__, "folder": folder}
	)
