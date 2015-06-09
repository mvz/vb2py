vb2Py - Code Modules
====================

Contents of this page:

- General_
- `Default Conversion`_      
- `List of Options`_

Options:

- TryToExtractDocStrings_


General
-------

Code modules are translated into Python code modules with a functions and subroutines being converted to module level functions. Variables, functions and subroutines which are declared as ``Public`` or ``Global`` are flagged as globally accessible which means that, if another module tries to access them, then that module will acquire an ``import`` statement to import the code module and the variable reference will be replaced by a fully qualified name (``module.name``).



Default Conversion
------------------

VB(VBCodeModule)::

    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub


List of Options
---------------

There are no specific options relating to code modules.

Some ``General`` options apply::

    [General]
	# Yes or No, whether to try to automatically extract docstrings from the code
	TryToExtractDocStrings = Yes


TryToExtractDocStrings
~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``TryToExtractDocStrings = Yes | No``

If ``TryToExtractDocStrings`` is set then any contiguous block of comment lines found at the start of the module are interpretted as a docstring and added to the class definition. The docstring terminates with the first non-comment line.

VB(VBCodeModule)::

    ' VB2PY-GlobalSet: General.TryToExtractDocStrings = Yes
	' This is the documentation for the module
	' This line is also documentation
	' So is this one
	' And this is the last

    Public Name As String
    Public Age As Single
    Private ID As Long

    Public Sub checkAge()
        If Age = 0 Then Age = 1
    End Sub
    '
    Private Sub setUp()
        ID = Rnd()
        If ID = 0 Then setUp
    End Sub
    ' VB2PY-Unset: General.TryToExtractDocStrings
