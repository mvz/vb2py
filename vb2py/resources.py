import os
from vb2py.utils import rootPath
from vb2py import logger
log = logger.getLogger("vb2Py")

twips_per_pixel = 15

class BaseResource(object):
    """A VB form resource object"""

    target_name = "Python"
    name = "baseResource"
    form_class_name = "FormClass"
    form_super_classes = []
    allow_new_style_class = 1

    def __init__(self, basesourcefile=None):
        """Initialize the resource"""
        if basesourcefile is None:
            self.basesourcefile = os.path.join(
                        rootPath(), "targets", self.target_name, "basesource")
        else:
            self.basesourcefile = basesourcefile
        #
        # Apply default resource
        self._rsc = {}
        self._code = ""
        #
        log.debug("BaseResource init")

    def updateFrom(self, form):
        """Update our resource from the form object"""
        #
        # The main properties of the form
        d = self._rsc['application']['backgrounds'][0]
        self.name = form.name
        d['name'] = form.name
        d['title'] = form.Caption

        #
        # Add menu height to form height if it is needed
        if form._getControlsOfType("Menu"):
            height_modifier = form.HeightModifier + form.MenuHeight
        else:
            height_modifier = form.HeightModifier

        d['size'] = (form.ClientWidth/twips_per_pixel, form.ClientHeight/twips_per_pixel+height_modifier)
        d['position'] = (form.ClientLeft/twips_per_pixel, form.ClientTop/twips_per_pixel)

        #
        # The components (controls) on the form
        c = self._rsc['application']['backgrounds'][0]['components']

        for cmp in form._getControlList():
            obj = form._get(cmp)
            entry = obj._getControlEntry()
            if entry:
                c += entry

        #
        # The menus
        m = []
        self._rsc['application']['backgrounds'][0]['menubar']['menus'] = m

        self.addMenus(form, m)

    def updateCode(self, form):
        """Update our code blocks"""
        #
        # Make sure the code structure has the right context
        form.code_structure.classname = self.form_class_name
        form.code_structure.superclasses = self.form_super_classes
        form.code_structure.allow_new_style_class = self.allow_new_style_class
        #
        # Convert it to Python code
        self.code_structure = form.code_structure

    def addMenus(self, obj, to_menu):
        """Add menus"""
        for mnu in obj._getControlsOfType("Menu"):
            d = mnu._pyCardMenuEntry()
            d["items"] = []
            to_menu.append(d)
            self.addMenus(mnu, d['items'])
            if not d['items']:
                del(d['items'])
                d['type'] = 'MenuItem'

    def writeToFile(self, basedir, write_code=0):
        """Write ourselves out to a directory"""
