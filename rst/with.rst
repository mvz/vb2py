vb2Py - With
============

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_

Specific options:

* EvaluateVariable_
* WithVariablePrefix_
* UseNumericIndex_


General
-------

In VB, ``With`` blocks represent a shorthand for refering to a specific object. Properties and methods of the object can be referenced by just prefixing the name with a '.' and the base object name is infered from the context. Python has no such equivalent and therefore the objects must be written out in full. By default, a ``with`` variable is created and this is used to fully qualify the references. This is *safe* in cases where evaluation of the ``With`` object is expensive [1]_ but may make the code look less clear. To get around this, an option is avaialble to use the actual variable name.


Default Conversion
------------------

VB::

    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With


List of Options
---------------

Here are the options in the INI file::

    [With]
    # Once or EachTime, how many times to evaluate the with variable
    EvaluateVariable = Once
    # Name of with variable (only used if EvaluateVariable is Once)
    WithVariablePrefix = _with
    # Yes or No, use numeric index on with variable (needed if you every have nested Withs and EvaluateVariable = Once)
    UseNumericIndex = Yes


EvaluateVariable
~~~~~~~~~~~~~~~~

Syntax: ``EvaluateVariable = Once | EveryTime``

The default behaviour is to evaluate the ``With`` object once at the start of the block. By setting this option to ``EachTime`` you can force the object to be evaluated each time it is required. This generally looks more natural but can lead to undesired side effects or slow run times depending on how expensive [1]_ the object is to calculate.

VB::

    ' VB2PY-Set: With.EvaluateVariable = EveryTime
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.EvaluateVariable


WithVariablePrefix
~~~~~~~~~~~~~~~~~~~~

Syntax: ``WithVariablePrefix = name``

When EvaluateVariable_ is set to ``Once``, this option determines the prefix used to name the variable used in the ``With``. If UseNumericIndex_ is set to ``No`` then this option sets the variable name used, otherwise this is the prefix and the final variable will also include a unique ID number.

VB::

    ' VB2PY-Set: With.WithVariablePrefix = withVariable
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.WithVariablePrefix



UseNumericIndex
~~~~~~~~~~~~~~~

Syntax: ``UseNumericIndex = Yes | No``

When EvaluateVariable_ is set to ``Once``, this option determines whether a unique ID number is appended to the WithVariablePrefix_ to determine the variable name used to hold the object. If used, the index is incremented for each ``With`` constuct found. This option is always required to be ``Yes`` where the code includes nested ``With`` blocks *and* EvaluateVariable_ is set to ``Once``. If neither of these conditions applies then it is safe to set this to ``No``

VB::

    ' VB2PY-Set: With.UseNumericIndex = No
    With MyObject
        .Height = 10
        .Width = .Height * .ScaleFactor
    End With
    ' VB2PY-Unset: With.UseNumericIndex


--------------

.. [1] Expensive as in CPU time.
