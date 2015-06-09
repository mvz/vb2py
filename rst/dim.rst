vb2Py - Dim and Redim
=====================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_


General
-------

VB variables can be defined or dimensioned using the ``Dim`` keyword to specify their type and/or size (for arrays). Python has no equivalent of the ``Dim`` statement but the converted code attempts to match the semantics of the ``Dim`` by,

1. Initializing the named variables to an empty object of the required type
2. Creating an array of the required size (using the ``vbfunctions.vbObjectInitialize`` function)

Note, however, that this conversion does not fix the type and size of the object and subsequent rebindings of the name will remove all reference to the previous *type*. Using the ``Preserve`` option with the ``ReDim`` command causes the original variable to be passed to the ``vbfunctions.vbObjectInitialize`` so that the contents of the object can be retained after resizing.


Default Conversion
------------------

VB::

    Dim a, b, c(20)
    Dim x As Integer, y As String, z(10) As Variant
    Dim obj As New Collection

    ReDim c(10)
    ReDim Preserve c(10)


List of Options
---------------

There are no options specific to the ``Dim`` statement.
