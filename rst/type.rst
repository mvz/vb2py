vb2Py - User Types
==================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_


General
-------

VB has the concept of a user ``Type``. A user can define a ``Type`` which can then be used to define variables which have one or more *properties*, each with its own type. The closest Python equivalent is a bare class. VB User ``Types`` are converted to bare classes with the appropriate properties. When a user ``Type`` is used, an istance of the class is created in the Python code.

Class properties are defined in an ``__init__`` method to avoid having mutable class properties.


Default Conversion
------------------

VB::

    Type Point
        X As Single
        Y As Single
    End Type
    '
    Type Line2D
        Start As Point
        Finish As Point
    End Type
    '
    Dim p1 As Point, p2 As Point
    p1.X = 10
    p1.Y = 20
    p2.X = 30
    p3.Y = 40
    '
    Dim l1 As Line2D
    l1.Start = p1
    l1.Finish = p2


List of Options
---------------

There are no options specific to the ``Type`` statement.
