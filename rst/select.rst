vb2Py - Select
==============

Contents of this page:

- General_
- `Default Conversion`_
- `List of Options`_

Specific options:

- EvaluateVariable_
- SelectVariablePrefix_
- UseNumericIndex_


General
-------

``Select`` blocks are replaced by ``if/elif/else`` blocks. By default a ``select`` variable
is created, which is used in subsequent tests. This means that the checked value is
only evaluated once per ``select`` and not once per condition. If this is not an issue then
an option allows the value to be evalutated each time as required. Also by default, a numeric index is appended to the select variable to prevent clashed for nested ``Select`` constructs.

The conversion handles multiple values per case and even range settings.


Default Conversion
------------------

VB::

    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select


List of Options
---------------

Here are the options in the INI file::

    [Select]
    # Once or EachTime, how many times to evaluate the case variable
    EvaluateVariable = Once
    # Name of select variable (only used if EvaluateVariable is Once)
    SelectVariablePrefix = _select
    # Yes or No, use numeric index on select variable (needed if you every have nested Selects and EvaluateVariable = Once)
    UseNumericIndex = Yes


EvaluateVariable
~~~~~~~~~~~~~~~~

Syntax: ``EvaluateVariable = Once | EachTime``

The default behaviour is to evaluate the select expression once at the start of the block. By setting this option to ``EachTime`` you can force the expression to be evaluated for each ``if/elif`` statement. This generally looks cleaner but can lead to undesired side effects or slow run times depending on how expensive [1]_ the expression is to calculate.

VB::

    ' VB2PY-Set: Select.EvaluateVariable = EachTime
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.EvaluateVariable


SelectVariablePrefix
~~~~~~~~~~~~~~~~~~~~

Syntax: ``SelectVariablePrefix = name``

When EvaluateVariable_ is set to ``Once``, this option determines the prefix used to name the variable used in the select. If UseNumericIndex_ is set to ``No`` then this option sets the variable name used, otherwise this is the prefix and the final variable will also include a unique ID number.

VB::

    ' VB2PY-Set: Select.SelectVariablePrefix = selectVariable
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.SelectVariablePrefix



UseNumericIndex
~~~~~~~~~~~~~~~

Syntax: ``UseNumericIndex = Yes | No``

When EvaluateVariable_ is set to ``Once``, this option determines whether a unique ID number is appended to the SelectVariablePrefix_ to determine the variable name used to hold the select expression. If used, the index is incremented for each ``select`` constuct found. This option is always required to be ``Yes`` where the code includes nested ``Select`` blocks *and* EvaluateVariable_ is set to ``Once``. If neither of these conditions applies then it is safe to set this to ``No``

VB::

    ' VB2PY-Set: Select.UseNumericIndex = No
    Select Case Value
        Case 1
            DoOne
        Case 2
            DoTwo
        Case 3, 4
            DoThreeOrFour
        Case 5 To 10
            DoFiveToTen
        Case Else
            DoElse
    End Select
    ' VB2PY-Unset: Select.UseNumericIndex

--------------

.. [1] Expensive as in CPU time.
