# Copyright (c) 2026, Sidharth PV and contributors
# See license.txt

import importlib
from pathlib import Path

import frappe
from frappe.tests.utils import FrappeTestCase

import fieldbook


class TestApp(FrappeTestCase):
	def test_display_title(self):
		self.assertEqual(frappe.get_hooks("app_title", app_name="fieldbook"), ["Fieldbook"])

	def test_every_listed_patch_exists(self):
		"""A line in patches.txt with no file behind it makes `bench migrate` fail on every site."""
		patches_txt = Path(fieldbook.__file__).parent / "patches.txt"
		listed = [
			line.split()[0]
			for line in patches_txt.read_text().splitlines()
			if line.strip() and not line.startswith(("#", "["))
		]
		for dotted in listed:
			module = importlib.import_module(dotted)
			self.assertTrue(callable(getattr(module, "execute", None)), dotted)

	def test_patches_are_readable_by_frappe(self):
		from frappe.modules.patch_handler import get_patches_from_app

		# must parse without error; this app currently lists no patches
		self.assertIsInstance(get_patches_from_app("fieldbook"), list)
