vb2Py - Do ... Loop
===================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_

Different forms:

* `Do ... Loop`_
* `Do While ... Loop`_
* `Do ... Loop While`_
* `Do Until ... Loop`_
* `Do ... Loop Until`_



General
-------

All variations of VB's ``Do ... Loop`` construct are converted to an equivalent Python ``while`` block. Preconditions are converted to the equivalent condition in the ``while`` statement itself, whereas post-conditions are implemented using an ``if ...: break`` . ``Exit's`` from the loop are also implemented using ``break`` . ``Until`` conditions (pre or post) are implemented by negating the condition itself but do not affect the structure.

Default Conversion
------------------

Do ... Loop
~~~~~~~~~~~

VB::

    Do
        Val = Val + 1
        If SomeCondition Then Exit Do
    Loop


Do While ... Loop
~~~~~~~~~~~~~~~~~

VB::

    Do While Condition
        Val = Val + 1
        If SomeCondition Then Exit Do
    Loop


Do ... Loop While
~~~~~~~~~~~~~~~~~

VB::

    Do
        Val = Val + 1
        If SomeCondition Then Exit Do
    Loop While Condition


Do Until ... Loop
~~~~~~~~~~~~~~~~~

VB::

    Do Until Condition
        Val = Val + 1
        If SomeCondition Then Exit Do
    Loop

Do ... Loop Until
~~~~~~~~~~~~~~~~~

VB::

    Do
        Val = Val + 1
        If SomeCondition Then Exit Do
    Loop Until Condition


List of Options
---------------

There are no options for the ``Do ... Loop`` construct.
