# Color Replace

Recursively searches through document to replace instances of desired color.

- Searches through any number of nested object groups.
- Easily choose color to find and color to replace by using selected object colors or by selecting with standard color-picker interface.
- Options for searching entire document or only selected objects.
- Options to replace fills and/or outline colors in found objects.

# How to Implement

1. Place nocolor2.jpg wherever you like.

2. Replace these lines with new file path reference:

```
findPreview.Picture = LoadPicture("S:\XI-Online\!-All Other Stuff\Ethan References\Macros\ColorReplace\nocolor2.jpg")
replacePreview.Picture = LoadPicture("S:\XI-Online\!-All Other Stuff\Ethan References\Macros\ColorReplace\nocolor2.jpg")
```

3. For default colors, replace any instances of current defaults, e.g. "Sublimation Black", "DTG White", "Black", "White"
and update with their RGB/CMYK values in findDrop_Change() and replaceDrop_Change()