

import NVDAObjects
import controlTypes
from logHandler import log
import config
import textInfos
import scriptHandler
from nvdaBuiltin.appModules.outlook import *

class DocumentSubMailNavigation(OutlookWordDocument):


    def isEmailDivision(self, div):
        try:
            return div.Range.Text.strip().lower().startswith("van:") or div.Range.Text.strip().lower().startswith("from:")
        except:
            return False



    @scriptHandler.script(
        gesture="kb:control+shift+."
    )
    def script_toNextSubMail(self, gesture):
        currentLocation = self.makeTextInfo("selection")._rangeObj.Start

        divisions = self.makeTextInfo("selection")._rangeObj.HTMLDivisions
        nextDivision = -1
        currentNextDivisionLocation = 10**20
        for ii in range(1, divisions.Count+1):
            if (divisions(ii).Range.Start > currentLocation):
                if (divisions(ii).Range.Start < currentNextDivisionLocation and self.isEmailDivision(divisions(ii)) ):
                    nextDivision = ii
                    currentNextDivisionLocation = divisions(ii).Range.Start

        if nextDivision > 0:
            divisions(nextDivision).Range.Select()
            ti = self.makeTextInfo("selection")
            log.info(ti.text)
            ti.collapse(end = True)
            ti.updateCaret()


    @scriptHandler.script(
        gesture="kb:control+shift+,"
    )
    def script_toPreviousSubMail(self, gesture):
        currentLocation = self.makeTextInfo("selection")._rangeObj.Start

        divisions = self.makeTextInfo("selection")._rangeObj.HTMLDivisions
        nextDivision = -1
        currentNextDivisionLocation = -10**20
        for ii in range(1, divisions.Count+1):
            if (divisions(ii).Range.Start < currentLocation):
                if (divisions(ii).Range.Start > currentNextDivisionLocation and self.isEmailDivision(divisions(ii))):
                    nextDivision = ii
                    currentNextDivisionLocation = divisions(ii).Range.Start


        if nextDivision > 0:

            if nextDivision > 1:
                while not self.isEmailDivision(divisions(nextDivision - 1)):
                    nextDivision -= 1
                    if nextDivision == 1:
                        break

            if nextDivision == 1:
                ti = self.makeTextInfo("first")
                ti.updateCaret()
                return


            divisions(nextDivision).Range.Select()
            ti = self.makeTextInfo("selection")
            ti.collapse(end = True)
            ti.updateCaret()

            # ti = self.makeTextInfo("first")
            # ti.move(textInfos.UNIT_CHARACTER,end )
            # ti.updateCaret()
