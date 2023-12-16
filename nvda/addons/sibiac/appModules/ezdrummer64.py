# -*- coding: utf-8 -*-
# Single Image Blob Interface Accessible Control
#
# EZ Drummer 2 64bit standalone overlay
#
# AZ (www.azslow.com), 2018 - 2020
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import appModuleHandler
import controlTypes
from NVDAObjects.window import Window
import core
import api

from .sibiac import *
from .sibiac.ezdrummer import EZDrummer

class EZDrummer_Standalone(Window):
	def _afterFocus(self):
		"""
		Move focus to reasonable child
		"""
		focus = api.getFocusObject()
		if self == focus:
			try:
				child = self.getChild(0)
				sibiac.MoveFocusTo(child.windowHandle)
			except:
				pass

	def event_gainFocus(self):
		""" Transfer focus to child """
		super(EZDrummer_Standalone,self).event_gainFocus()
		# give some time, then check (beter way?)
		core.callLater(0.5, self._afterFocus)

	

class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if obj.windowClassName == "wxWindowNR" and obj.role == controlTypes.ROLE_PANE :
			clsList.insert(0, EZDrummer_Standalone)
		elif obj.windowClassName.startswith("Plugin"):
			clsList.insert(0, EZDrummer)
