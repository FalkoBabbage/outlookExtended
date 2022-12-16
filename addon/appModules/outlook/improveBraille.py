

import NVDAObjects
import controlTypes
from logHandler import log
import config
from nvdaBuiltin.appModules.outlook import *

executedVerbLabels={
	# Translators: the last action taken on an Outlook mail message
	VERB_REPLYTOSENDER:_("r"),
	# Translators: the last action taken on an Outlook mail message
	VERB_REPLYTOALL:_("ra"),
	# Translators: the last action taken on an Outlook mail message
	VERB_FORWARD:_("f"),
}

class brailleAbbriviations(UIAGridRow):
	def __init__(self, *args, **kwargs):
		self.brailleName = True
		super().__init__(*args, **kwargs)

	def _cache_name(self):
		return False
	def _get_name(self):
		try:
			if self.brailleName:
				return self.getBrailleName()
		except:
			self.brailleName = True
		return super()._get_name()

	def getBrailleName(self):
		textList = []
		if controlTypes.State.EXPANDED in self.states:
			textList.append(controlTypes.State.EXPANDED.displayString)
		elif controlTypes.State.COLLAPSED in self.states:
			textList.append(controlTypes.State.COLLAPSED.displayString)
		selection = None
		if self.appModule.nativeOm:
			try:
				selection = self.appModule.nativeOm.activeExplorer().selection.item(1)
			except COMError:
				pass
		if selection:
			try:
				unread = selection.unread
			except COMError:
				unread = False
			# Translators: when an email is unread
			if unread: textList.append(_("u"))
			try:
				mapiObject = selection.mapiObject
			except COMError:
				mapiObject = None
			if mapiObject:
				v = comtypes.automation.VARIANT()
				res = NVDAHelper.localLib.nvdaInProcUtils_outlook_getMAPIProp(
					self.appModule.helperLocalBindingHandle,
					self.windowThreadID,
					mapiObject,
					PR_LAST_VERB_EXECUTED,
					ctypes.byref(v)
				)
				if res == S_OK:
					verbLabel = executedVerbLabels.get(v.value, None)
					if verbLabel:
						textList.append(verbLabel)
			try:
				attachmentCount = selection.attachments.count
			except COMError:
				attachmentCount = 0
			# Translators: when an email has attachments
			if attachmentCount > 0: textList.append(_("a"))
			try:
				importance = selection.importance
			except COMError:
				importance = 1
			importanceLabel = importanceLabels.get(importance)
			if importanceLabel: textList.append(importanceLabel)
			try:
				messageClass = selection.messageClass
			except COMError:
				messageClass = None
			if messageClass == "IPM.Schedule.Meeting.Request":
				# Translators: the email is a meeting request
				textList.append(_("mr"))
		childrenCacheRequest = UIAHandler.handler.baseCacheRequest.clone()
		childrenCacheRequest.addProperty(UIAHandler.UIA_NamePropertyId)
		childrenCacheRequest.addProperty(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId)
		childrenCacheRequest.TreeScope = UIAHandler.TreeScope_Children
		# We must filter the children for just text and image elements otherwise getCachedChildren fails completely in conversation view.
		childrenCacheRequest.treeFilter = createUIAMultiPropertyCondition({UIAHandler.UIA_ControlTypePropertyId: [
			UIAHandler.UIA_TextControlTypeId, UIAHandler.UIA_ImageControlTypeId]})
		cachedChildren = self.UIAElement.buildUpdatedCache(childrenCacheRequest).getCachedChildren()
		if not cachedChildren:
			# There are no children
			# This is unexpected here.
			log.debugWarning("Unable to get relevant children for UIAGridRow", stack_info=True)
			return super(UIAGridRow, self).name
		for index in range(cachedChildren.length):
			e = cachedChildren.getElement(index)
			UIAControlType = e.cachedControlType
			UIAClassName = e.cachedClassName
			# We only want to include particular children.
			# We only include the flagField if the object model's flagIcon or flagStatus is set.
			# Stops us from reporting "unflagged" which is too verbose.
			if selection and UIAClassName == "FlagField":
				try:
					if not selection.flagIcon and not selection.flagStatus: continue
				except COMError:
					continue
			# the category field should only be reported if the objectModel's categories property actually contains a valid string.
			# Stops us from reporting "no categories" which is too verbose.
			elif selection and UIAClassName == "CategoryField":
				try:
					if not selection.categories: continue
				except COMError:
					continue
			# And we don't care about anything else that is not a text element.
			elif UIAControlType != UIAHandler.UIA_TextControlTypeId:
				continue
			name = e.cachedName
			columnHeaderTextList = []
			if name and config.conf['documentFormatting']['reportTableHeaders']:
				columnHeaderItems = e.getCachedPropertyValueEx(UIAHandler.UIA_TableItemColumnHeaderItemsPropertyId,
															   True)
			else:
				columnHeaderItems = None
			if columnHeaderItems:
				columnHeaderItems = columnHeaderItems.QueryInterface(UIAHandler.IUIAutomationElementArray)
				for index in range(columnHeaderItems.length):
					columnHeaderItem = columnHeaderItems.getElement(index)
					columnHeaderTextList.append(columnHeaderItem.currentName)
			columnHeaderText = " ".join(columnHeaderTextList)
			if columnHeaderText:
				text = u"{name}".format( name=name)
			else:
				text = name
			if text:
				if UIAClassName == "FlagField":
					pass
					textList.insert(0, text)
				else:
					text += u","
					textList.append(text)
		return " ".join(textList)





	def reportFocus(self):
		self.brailleName = False
		self.name = super()._get_name()
		ret = super().reportFocus()
		self.brailleName = True
		self.name = self.getBrailleName()
		return ret