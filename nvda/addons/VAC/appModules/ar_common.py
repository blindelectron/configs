"""Name Virtual Audio Cable app fields by controlID.

Author: Doug Lee
"""

import appModuleHandler
import controlTypes
import api
from NVDAObjects.behaviors import ProgressBar
import time
import sys
from wx import CallLater

import languageHandler, addonHandler
addonHandler.initTranslation()

class AppModule(appModuleHandler.AppModule):
	# Tab/ShiftTab trigger moving mouse to current position, which causes speaking of help balloons for fields.
	__gestures = {
		"kb:tab": "makeMouseFollow",
		"kb:shift+tab": "makeMouseFollow"
	}
	def script_makeMouseFollow(self, gesture):
		gesture.send()
		CallLater(200, self._mouseToFocus)

	def _mouseToFocus(self):
		"""Move the mouse to focus to cause the current field's help balloon to appear and speak.
		"""
		api.moveMouseToNVDAObject(api.getFocusObject())

	def event_NVDAObject_init(self,obj):
		if obj.role == controlTypes.ROLE_HELPBALLOON:
			# No need to say "balloon" when these pop up on every focus change.
			obj.roleText = " "
		elif obj.role not in [controlTypes.ROLE_BUTTON, controlTypes.ROLE_LIST, controlTypes.ROLE_LISTITEM]:
			# These tend to be wrong; let the help balloon name the field in a moment.
			# Examples caught: ComboBox, Checkbox, Edit.
			obj.name = ""

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		# Stop progress bars from speaking, as they are used as sound level meters and such.
		if obj.windowClassName=="msctls_progress32":
			try:
				clsList.remove(ProgressBar)
			except ValueError:
				pass

