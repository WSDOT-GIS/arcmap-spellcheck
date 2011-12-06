import sys, os, subprocess

# Get the directory that Inkscape is in.
inkscape = os.path.join(os.path.join(os.environ["programfiles"], "Inkscape"), "Inkscape.exe")

# Exit the program if Inkscape cannot be found.
if not os.path.exists(inkscape):
    exit("Inkscape executable file not found.")

svg = os.path.abspath("SpellCheckIcon.svg")
buttonPng = os.path.abspath("SpellCheckButton.png")
addInPng = os.path.abspath("SpellCheckAddIn.png")

def exportPng(svgFile, pngFile, size):
    if os.path.exists(pngFile):
        print 'Output file "%s" already exists.  Skipping to next SVG file...' % pngFile
    else:
        print 'Exporting "%s" from "%s"...' % (pngFile, svgFile)
        returncode = subprocess.call([inkscape, svgFile, "--export-png=%s" % pngFile, "-w%s" % size, "-h%s" % size, "--export-area-drawing"])
        print "Return code: %s" % returncode
        return returncode

if not os.path.exists(svg):
    # Print an error message if the SVG file was not found.
    print '"%s" was not found." % svgFile'
else:
    exportPng(svg, buttonPng, 128)
    exportPng(svg, addInPng, 64)
