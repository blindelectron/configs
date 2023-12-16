# Single Image Blob Interface Accessible Control
# 
# sforzando overlay
#
# AZ (www.azslow.com), 2018
from logHandler import log

from . import *


class Sforzando(SIBINVDA):

	sname = "Sforzando"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 774, 498)
		self.sibi.add( Clickable("Instrument", self.sibi, Pt(88, 21), Pt(232, 36), ("dynamic_text", "silent_action")) )
		self.sibi.add( Clickable("Polyphony", self.sibi, Pt(488, 43), Pt(520, 56), ("dynamic_text", "silent_action")) )
		self.sibi.add( Clickable("Pitchbend range", self.sibi, Pt(571, 43), Pt(603, 56), ("dynamic_text", "silent_action")) )
