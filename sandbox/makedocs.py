import vb2py.vbparser
import re
import glob
import vb2py.utils
import os
import sys

from docutils.core import publish_string

#
# Turn off logging
import vb2py.extensions
vb2py.extensions.disableLogging()
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff


def doAutomaticVBConversion(txt):
    """Convert VB code in the text to Python"""
    vb = re.compile(r'(.*?)<p>VB(\(.*?\))?:</p>.*?<pre class="literal-block">(.*?)</pre>(.*?)',
                    re.DOTALL+re.MULTILINE)
    def convertVB(match):
        """Convert the match"""
        if match.groups()[1]:
            mod = getattr(vb2py.vbparser, match.groups()[1][1:-1])()
        else:
            mod = vb2py.vbparser.VBCodeModule()
        mod.importStatements = lambda x : ""
        m = vb2py.vbparser.parseVB(match.groups()[2], container=mod)
        return '%s<table style="code-table">' \
               '<tr><th class="code-header">VB</th><th class="code-header">Python</th></tr>' \
               '<tr><td class="vb-code-cell"><pre>%s</pre></td>' \
               '<td class="python-code-cell"><pre>%s</pre></td></tr></table>%s' % (
                                           match.groups()[0],
                                           match.groups()[2].replace("\n", "<br>"),
                                           m.renderAsCode(1).replace("\n", "<br>"),
                                           match.groups()[3])
    return vb.sub(convertVB, txt)


def addToTemplate(text, template, token):
    """Add the text to the template"""
    template_text = open(template, "r").read()
    template_text = template_text.replace(token, text)
    return template_text


if __name__ == "__main__":

    add_to_template = 0
    if len(sys.argv) == 1:
        files = glob.glob(os.path.join("rst", "*.rst"))
    else:
        files = sys.argv[1:]
    settings = {
            'embed_stylesheet' : False,
            'stylesheet' : 'default.css',
            'stylesheet_path' : '',
            }

    print "\nvb2Py documentation generator\n"
    for fn in files:
        print "Processing '%s' ..." % fn,

        if not fn.startswith("rst/"):
            print "Source must be in rst/"
            break
        if not fn.endswith(".rst"):
            print "Source must be reStructuredText"
            break

        target_file = "doc/" + fn[4:-4] + ".html"

        print "to %s ..." % target_file

        txt = open(fn, "r").read()

        base_html_text = publish_string(writer_name = 'html',
                source = txt,
                settings_overrides = settings)

        marked_up_text = doAutomaticVBConversion(base_html_text)

        if add_to_template:
            marked_up_text = addToTemplate(marked_up_text, sys.argv[2], "DOCSGOHERE")

        open(target_file, "w").write(marked_up_text)

    print "Done!"
