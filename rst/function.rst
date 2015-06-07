vb2Py - Functions
=================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_

Specific options:

* ReturnVariableName_
* PreInitializeReturnVariable_

Additional

* `Missing Arguments`_
* `Argument Passing`_

General
-------

Functions are converted to Python functions with an explicit return statement. By default, a return variable is created and initialize as soon as the function starts (in case the function exits before the variable is assigned a proper value). Assignments to the VB function name are mapped to the return variable.

Local variables in the VB functions are also local in the Python version. If a module global is used on the left hand side of an assignment then a Python ``global`` statement will be inserted at the head of the function. Project *globals* will be replaced by their fully-qualified versions.


Default Conversion
------------------

VB::

    Dim moduleGlobal1, moduleGlobal2

    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)


Missing Arguments
~~~~~~~~~~~~~~~~~

Optional arguments which are not supplied and have no defaults are initialized with the ``VBMissingArgument`` object. This object can be detected by the ``vbfunctions.IsMissing`` function to provide initialization of missing arguments within the body of the function. This functionality is transparent under normal conditions, but if the function manually assigns a value to the missing parameter prior to the IsMissing call then the behaviour may not match that of VB, since the ``vbfunctions.IsMissing`` function has no way to detect that the parameter was not supplied.

VB::

    Dim moduleGlobal1, moduleGlobal2

    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        If IsMissing(Y) Then Y = 12
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)


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

Here are the options in the INI file::

    [Functions]
    # Name of variable used in Functions
    ReturnVariableName = _ret
    # Yes or No, leave at Yes unless good reasons!
    PreInitializeReturnVariable = Yes


ReturnVariableName
~~~~~~~~~~~~~~~~~~

Syntax: ``ReturnVariableName = name``

This option allows the return variable name to be specified. No checking is done to ensure that the name does not clash with local or global variables, so care should be taken when selecting a suitable name.

VB::

    Dim moduleGlobal1, moduleGlobal2

    ' VB2PY-GlobalSet: Functions.ReturnVariableName = _MyFunc
    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function
    ' VB2PY-Unset: Functions.ReturnVariableName

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)


PreInitializeReturnVariable
~~~~~~~~~~~~~~~~~~~~~~~~~~~

Syntax: ``PreInitializeReturnVariable = Yes | No``

By default the return variable is initialized to ``None`` at the start of the function so that an error does not occur in the event that the function returns before the return variable has been assigned to. This option allows this initialization step to be omitted and is safe as long as all return paths from the function include an explicit assignment to the return value variable.

VB::

    Dim moduleGlobal1, moduleGlobal2

    ' VB2PY-GlobalSet: Functions.PreInitializeReturnVariable = No
    Function MyFunc(X, Optional Y, Optional Z=20)
        Dim subLocal
        subLocal = X + Y + Z + moduleGlobal
        moduleGlobal2 = moduleGlobal2 + 1
        MyFunc = subLocal*10
    End Function
    ' VB2PY-Unset: Functions.PreInitializeReturnVariable

    a = MyFunc(1, 2)
    a = MyFunc(1, Z:=10)
