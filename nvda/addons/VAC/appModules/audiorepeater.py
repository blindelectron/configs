#Addon audiorepeater.py for NonVisual Desktop Access (NVDA)
#Copyright (C) 2014-2023 Doug Lee
#This file is covered by the GNU General Public License.
#See the file COPYING for more details.

import sys
if sys.version_info.major < 3:
	from ar_common import AppModule
else:
	from .ar_common import AppModule

