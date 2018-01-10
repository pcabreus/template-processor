TemplateProcessor
=================

This library add some extra functionalities to PhpWord library

Some functions:
* Added ``setImage`` function to work with images, mainly PositibeMedia images.
* Added two special characters, the ex (TemplateProcessor::SIGN_EX) and empty square (TemplateProcessor::SIGN_BLOCK).
* Added ``setValueBreakLine`` function to force to write a new line text.

**Warning**: These functions may not work properly on some Word Reader like (Libre Office or old Microsoft Office Word)