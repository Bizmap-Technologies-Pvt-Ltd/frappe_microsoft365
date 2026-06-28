"""Row-level scoping: a user sees only their own Microsoft Calendar; System Manager sees all."""

import frappe


def _is_admin(user):
	return user == "Administrator" or "System Manager" in frappe.get_roles(user)


def calendar_pqc(user=None):
	user = user or frappe.session.user
	if _is_admin(user):
		return ""
	return f"`tabMicrosoft Calendar`.`user` = {frappe.db.escape(user)}"


def calendar_has_permission(doc, user=None, ptype="read"):
	user = user or frappe.session.user
	if _is_admin(user):
		return True
	return (doc.user or user) == user
