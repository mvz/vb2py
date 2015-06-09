vb2Py - Subroutines
===================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_

Additional

* `Missing Arguments`_
* `Argument passing`_

General
-------

Subroutines are converted to Python functions with no return statement. Local variables in the VB subroutine are also local in the Python version. If a module global is used on the left hand side of an assignment then a Python ``global`` statement will be inserted at the head of the function. Project *globals* will be replaced by their fully-qualified versions.


Default Conversion
------------------

VB::

    Dim moduleGlobal1, moduleGlobal2

    Sub MySub(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
    End Sub

    MySub 1, 2
    MySub 1, Z:=10


Missing Arguments
~~~~~~~~~~~~~~~~~

Optional arguments which are not supplied and have no defaults are initialized with the ``VBMissingArgument`` object. This object can be detected by the ``vbfunctions.IsMissing`` function to provide initialization of missing arguments within the body of the subroutine. This functionality is transparent under normal conditions, but if the subroutine manually assigns a value to the missing parameter prior to the IsMissing call then the behaviour may not match that of VB, since the ``vbfunctions.IsMissing`` function has no way to detect that the parameter was not supplied.

VB::

    Dim moduleGlobal1, moduleGlobal2

    Sub MySub(X, Optional Y, Optional Z=20)
        Dim subLocal
        If IsMissing(Y) Then Y = 12
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
    End Sub

    MySub 1, 2
    MySub 1, Z:=10


Argument Passing
~~~~~~~~~~~~~~~~

VB has two argument passing schemes,

1. ``ByRef`` (Default) - arguments are passed by reference. Changes to the value inside the
   subroutine are reflected in the corresponding parameter in the calling scope.
2. ``ByVal`` - arguments are passed by value. Changes to the value inside the subroutine do
   not affect the parameter in the calling scope.

Although Python's argument passing semantics are often refered to as pass-by-reference, the actual behaviour does not always match VB's ``ByRef`` because of immutable object types and name re-binding. Although there are technical solutions for these issues, the current version of vb2Py does not make any attempt to match behaviours.

**This means that the following code, although converted, does not behave the same in the Python version**

VB::

    Sub DoIt(x, ByVal y)
        x = x + 1
        y = y + 1
    End Sub

    x = 0
    y = 0
    DoIt x, y
    ' x is now 1, y is still 0


List of Options
---------------

There are no options specific to the ``Sub`` statement.
