vb2Py - Enumerations
====================

Contents of this page:

* General_
* `Default Conversion`_
* `List of Options`_


General
-------

``Enumeration`` types in VB allow the representation of a single value by a number of different *states*. Python has no equivalent concept and so enumerations are simply mapped to a series of integer values. The enumeration names become local or global names depending on the scope of the enumeration. If the values for the enumeration items are specified then these are used, if not then the values are chosen sequentially starting at zero.


Default Conversion
------------------

VB::

    Enum Number
		One = 1
		Two = 2
		Three = 3
		Four = 4
		Ten = 10
		Hundred = 100
	End Enum

	Enum Day
		Mon
		Tue
		Wed
		Thu
		Fri
	End Enum



List of Options
---------------

There are no options specific to the ``Enum`` statement.
