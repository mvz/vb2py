vb2Py - For and For Each
========================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_

Different forms:

* `For i = 0 To 10`_
* `For Each Obj In Container`_



General
-------

``For`` and ``For Each`` statements are converted to an equivalent Python for block. Where the iteration is over an iterable object, the translation just uses the iterable. Where the VB statement is an iteration between two numbers, the ``vb2Py`` function ``vbForRange`` is used to match the behaviour.


Default Conversion
------------------

For i = 0 To 10
~~~~~~~~~~~~~~~

VB::

    For i = 0 To 10
        DoSomething i
        If Condition Then Exit For
    Next i


For Each Obj In Container
~~~~~~~~~~~~~~~~~~~~~~~~~

VB::

    For Each Obj In Container
        DoSomethingWith i
        If i.Condition Then Exit For
    Next i


List of Options
---------------

There are no options specific to the ``For`` statement.
