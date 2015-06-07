vb2Py - #If Directives
======================

*THE #IF DIRECTIVE DOES NOT WORK CORRECTLY IN v0.2*

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_


General
-------

``#If`` directives in VB are used to produce compile time switches (eg to define different constants or functions). Although Python doesn't have this concept, any of the use cases for #If directives are easily handled by Python's conventional ``if`` statement at runtime. Consequently, the ``#If`` conversion is simply converted to the equivalant Python ``if``.


Default Conversion
------------------

VB::

    #If Value = 10 Or Value = 20 Then
        Const MaxCols = 10
    #ElseIf Value = 30 Then
        Const MaxCols = 20
    #Else
        Const MaxCols = 30
    #End If


List of Options
---------------

There are no options specific to the ``If`` statement.
