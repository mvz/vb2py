vb2Py - Form Modules
====================

Contents of this page:

- General_
- `Default Conversion`_
- `List of Options`_

Options:

- UseNewStyleClasses_
- RespectPrivateStatus_
- PrivateDataPrefix_
- TryToExtractDocStrings_


General
-------

Form modules are translated into Python code modules with a single class whose name is the nane of the VB form. By default, this class is a *new style* Python class (inherits from ``Object``). All methods in the class are converted to unbound methods of the class. Properties are converted to Python properties but an error is raised if the property has both ``Let`` and ``Set`` decorators. Since Python has no equivalent of the ``Set`` keyword, the ``Property Set`` method is treated in the same way as a ``Property Let``.

Attributes defined at the form level are assumed to be class attributes in the Python class. By default, the conversion respects the Public/Private scope of both attributes and methods but this can be disabled if desired.


Default Conversion
------------------

VB(VBFormModule)::

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

Here are the options in the INI file::

    [Classes]
    # Yes or No, whether to use new style classes for all classes
    UseNewStyleClasses = Yes


In addition to these specific options, some ``General`` options apply::

    [General]
    # Yes or No, whether to respect Private status of variables
    RespectPrivateStatus = Yes
    # Prefix to use to tag data as private (Python normally uses __ but VB convention is m)
    PrivateDataPrefix = __
	# Yes or No, whether to try to automatically extract docstrings from the code
	TryToExtractDocStrings = Yes


UseNewStyleClasses
~~~~~~~~~~~~~~~~~~

By default, all forms are created as *new style* Python classes (inheriting from ``Object``). Old style classes can be created by setting the ``UseNewStyleClasses`` option to ``No``.

VB(VBFormModule)::

    ' VB2PY-GlobalSet: Classes.UseNewStyleClasses = No
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
    ' VB2PY-Unset: Classes.UseNewStyleClasses


RespectPrivateStatus
~~~~~~~~~~~~~~~~~~~~

Syntax: ``RespectPrivateStatus = Yes | No``

By default, variables or methods defined as Private (which is the default in VB), will be marked as private in the Python module also. Private Python variables will be prefixed with a private marker (two underscores by default). Since ``Private`` is the default in VB, this can lead to a lot of hidden variables in the Python code. The ``RespectPrivateStatus`` option allows you to turn off the ``Private/Public`` switch.

VB(VBFormModule)::

    ' VB2PY-GlobalSet: General.RespectPrivateStatus = No
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
    ' VB2PY-Unset: General.RespectPrivateStatus


PrivateDataPrefix
~~~~~~~~~~~~~~~~~

Syntax: ``PrivateDataPrefix = prefix``

If ``RespectPrivateStatus`` is set then each ``Private`` variable will be prefixed with the string specified by the ``PrivateDataPrefix`` option. By default this is two underscores, ``__``, which means that Python will use *name mangling* to ensure that the names really are private. Changing this option allows names to converted to some other convention (eg ``m``) which marks names but does not enforce privacy.

VB(VBFormModule)::

    ' VB2PYGlobal-Set: General.PrivateDataPrefix = m
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


TryToExtractDocStrings
~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``TryToExtractDocStrings = Yes | No``

If ``TryToExtractDocStrings`` is set then any contiguous block of comment lines found at the start of the module are interpretted as a docstring and added to the class definition. The docstring terminates with the first non-comment line.

VB(VBFormModule)::

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
