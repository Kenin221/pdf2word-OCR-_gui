from fontTools.ttLib import TTFont
font = TTFont('setofont.ttf')
font.save('./',reorderTables=True)