import wx

import addonHandler
import gui
from .donate_dialog import requestDonations
from . import patterns
import synthDriverHandler
import tones
addonHandler.initTranslation()

class TTAnnouncerSettingsPanel(gui.SettingsPanel):
    title = _("tt-announcer")

    def makeSettings(self, settingsSizer):
        sizer = gui.guiHelper.BoxSizerHelper(self, sizer=settingsSizer)
        synthEngine = _('&Synthesizer engine to use:')
        drivers = synthDriverHandler.getSynthList()
        self.synthNames = [synth[0] for synth in drivers if synth[0] != synthDriverHandler.getSynth().name]
        options = [synth[1] for synth in drivers if synth[0] != synthDriverHandler.getSynth().name]
        self.synthEngineCB = sizer.addLabeledControl(synthEngine, wx.Choice, choices=options)
        try:
            self.synthEngineCB.SetSelection(self.synthNames.index(self.addonConf["engine"]))
        except ValueError:
            self.synthEngineCB.SetSelection(0)
        self.regexTranslationCHK = sizer.addItem(wx.CheckBox(self, label=_("Consider  NVDA locale for regular expressions")))
        self.regexTranslationCHK.SetValue(self.addonConf['regexTranslation'])
        self.eventsList = sizer.addLabeledControl("events", gui.nvdaControls.CustomCheckListBox, choices=[])
        self.donateBtn = sizer.addItem(wx.Button(self, label=_("Support an author...")))
        self.donateBtn.Bind(wx.EVT_BUTTON, self.onDonate)
        self.fillEventsList()

    def onDonate(self, evt):
        requestDonations(self)

    def fillEventsList(self):
        checkedItems = []
        for index, event in enumerate(patterns.events):
            self.eventsList.Append(event["label"], event["name"])
            if event["name"] in self.addonConf["events"]:
                checkedItems.append(index)
        self.eventsList.CheckedItems = checkedItems
        self.eventsList.Select(0)

    def postInit(self):
        self.synthEngineCB.SetFocus()

    def isValid(self):
        if not self.eventsList.CheckedItems:
            gui.messageBox(_("At least one event must be selected to make addon function"), _("Error"), wx.OK|wx.ICON_ERROR,self)
            return False
        return super().isValid()

    def onSave(self):
        self.addonConf["engine"] = self.synthNames[self.synthEngineCB.GetSelection()]
        regexTranslation = self.addonConf["regexTranslation"]
        self.addonConf["regexTranslation"] = self.regexTranslationCHK.GetValue()
        self.addonConf["events"] = [self.eventsList.GetClientData(index) for index in self.eventsList.CheckedItems]
        self.onSaveCallback(self.addonConf["regexTranslation"] != regexTranslation)
